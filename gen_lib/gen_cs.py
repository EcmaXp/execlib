import io
import keyword
import re
from contextlib import contextmanager
from enum import IntEnum
from functools import lru_cache
from pathlib import Path
from typing import Set, Tuple, IO, Union, Dict, Match, Optional

import dataclasses
from dataclasses import dataclass

RE_ANNOTATION_SINGLE = re.compile(r"\[(.*?)\]")
RE_ANNOTATION = re.compile(r"^\[(.+?)\]$")
RE_FUNCTION = re.compile(r"^(\S+?) "  # type
                         r"(\S+?)"  # name
                         r"\("
                         r"(.*?)"
                         r"\)"
                         r";$"  # semi
                         )

RE_PROPERTY = re.compile(r"^(\S+?) "  # type
                         r"(\S+?) "  # name
                         r"\{"
                         r"\s*"
                         r"(.*?)"
                         r"\s*"
                         r"\}"
                         r"$"  # semi
                         )

RE_BRAKET = re.compile(r"^([{}])$")

CS_DEFAULT_TYPES = frozenset({
    "object",
    "string",
    "int",
    "bool",
    "void",
    "double",
    "float",
})

CS_KEYWORDS = CS_DEFAULT_TYPES | frozenset({
    "fixed",
    "volatile",
    "goto",
    "checked",
    "operator",
    "namespace",
    "base",
    "decimal",
    "default",
})


@dataclass
class Argument:
    type: str
    name: str
    io_type: str = None
    deco: str = None
    default: str = None


@dataclass
class Function:
    type: str
    name: str
    args: Tuple[Argument]
    py_args: Tuple[Argument] = None


@dataclass
class Property:
    type: str
    name: str
    accessor: Set[str]


@dataclass
class Class:
    name: str
    members: Dict[str, Union[Function, Property]]
    namespace: str
    header: str


@dataclass
class Enum:
    name: str
    values: Dict[str, int]
    namespace: str
    header: str


class OptinalType(IntEnum):
    HasDefault = 1
    HasDefine = 2


def get_optinal_type(arg: Argument) -> Optional[OptinalType]:
    if arg.default is not None:
        return OptinalType.HasDefault
    elif arg.deco is not None and "Optional" in arg.deco:
        return OptinalType.HasDefine
    else:
        return None


def camel_case_split(identifier):
    matches = re.finditer('.+?(?:(?<=[a-z])(?=[A-Z])|(?<=[A-Z])(?=[A-Z][a-z])|$)', identifier)
    return [m.group(0) for m in matches]


def camel_to_py_name(name, is_function):
    key = "_".join(map(str.lower, camel_case_split(name)))
    if keyword.iskeyword(key):
        key += "_"

    if not is_function and key in CS_KEYWORDS:
        key = name

    return key


def escape_cs_func_name(name):
    return name if name not in CS_KEYWORDS else f"@{name}"


CLASS_TREE = {
    "Application": "Application",
    "Workbooks": "Application",
    "Workbook": "Workbook",
    "Sheets": "Workbook",
    "Worksheet": "Sheets",
    "Range": "Worksheet",
}


class Parser:
    def __init__(self, name: str, finder: "ClassFinder"):
        self.name = name
        self.finder = finder

    def parse(self):
        name = self.name
        code = self.finder.read_code(self.name)

        namespace, header, lines = self.parse_headers(code)
        if "interface" in header:
            members = {node.name: node for node in self.parse_class_body(lines)}
            cls = Class(name, members, namespace, header)
            cls = self.patch_cls(cls)
            return cls
        elif "enum" in header:
            members = dict(self.parse_enum_body(lines))
            enum = Enum(name, members, namespace, header)
            return enum
        else:
            raise Exception(header)

    @staticmethod
    def parse_headers(code):
        lines = code.splitlines()
        count = 0
        namespace = None
        header = None
        for no, line in enumerate(lines):
            line = line.strip()
            if line.startswith("\ufeff//"):
                continue
            elif line.startswith("//"):
                continue
            elif line.startswith("["):
                continue
            elif line.startswith("using"):
                continue
            elif line == "{":
                count += 1
                if count == 2:
                    break
            elif line.startswith("namespace"):
                namespace = line.partition(" ")[-1]
            else:
                header = line
        else:
            assert False

        return namespace, header, lines[no:]

    @staticmethod
    def parse_class_body(lines):
        for line in lines:
            line = line.strip()
            anno = RE_ANNOTATION.match(line)
            if anno:
                continue

            if "this[]" in line:
                continue

            m_func = RE_FUNCTION.match(line)
            if m_func:
                deco_map = {}

                def register_deco(s: Match):
                    index = f"#{len(deco_map) + 1}"
                    deco_map[index] = s[1]
                    return index

                args = RE_ANNOTATION_SINGLE.sub(register_deco, m_func[3])
                args = args.split(",")
                new_args = []
                for arg in args:
                    arg = arg.strip()
                    if not arg:
                        assert len(args) == 1
                        break

                    tokens = arg.split()
                    assert len(tokens) >= 2, arg
                    default = None
                    if tokens[-2] == "=":
                        default = tokens.pop()
                        sep = tokens.pop()
                        assert sep == "="

                    assert len(tokens) >= 2, arg
                    name = tokens.pop()
                    type = tokens.pop()
                    token = tokens.pop()
                    if token in deco_map:
                        io_type, deco_index = None, token
                    else:
                        io_type, deco_index = token, tokens.pop() if tokens else None

                    assert not tokens, (tokens, line)

                    deco = deco_map[deco_index] if deco_index is not None else None
                    arg = Argument(type, name, io_type, deco, default)

                    new_args.append(arg)

                func = Function(m_func[1], m_func[2], tuple(new_args))
                yield func

            m_prop = RE_PROPERTY.match(line)
            if m_prop:
                new_accessor = set()
                for accessor in m_prop[3].split(";"):
                    accessor = accessor.strip()
                    if not accessor:
                        continue

                    accessor = accessor.rpartition(" ")[-1]
                    assert accessor in {"get", "set"}, repr(accessor)
                    new_accessor.add(accessor)

                yield Property(m_prop[1], m_prop[2], new_accessor)

    @staticmethod
    def parse_enum_body(lines):
        num = 0
        for line in lines:
            line = line.strip()
            anno = RE_ANNOTATION.match(line)
            if anno:
                continue

            if RE_BRAKET.match(line):
                continue

            line, sep, comment = line.partition("//")
            line = line.strip()
            if not line.endswith(","):
                print(repr(line))
                continue

            line = line.rstrip(",")
            key, sep, value = line.partition("=")
            key = key.strip()
            value = value.strip()
            if value:
                num = int(value)

            yield key, num
            num += 1

    @staticmethod
    def patch_cls(cls: Class):
        ws_type = "Worksheet"

        # noinspection PyUnusedLocal
        func: Function
        # noinspection PyUnusedLocal
        prop: Property

        if cls.name in CLASS_TREE:
            prop = cls.members["Parent"]
            prop.type = CLASS_TREE[cls.name]

        if cls.name == "Application":
            prop = cls.members["Selection"]
            prop.type = "Range"
        elif cls.name == "Sheets":
            func = cls.members["get_Item"]
            func.type = ws_type
            for arg in func.args:
                arg.default = arg.deco = None

            func = cls.members["Add"]
            func.args[0].type = func.args[1].type = ws_type
            func.args[2].type = "int"

            func = cls.members["Copy"]
            func.args[0].type = func.args[1].type = ws_type

            func = cls.members["Move"]
            func.args[0].type = func.args[1].type = ws_type

            func = cls.members["Select"]
            func.args[0].type = ws_type
        elif cls.name == "Range":
            func = cls.members["get_Item"]
            func.type = "Range"
            func.args[0].type = func.args[1].type = "int"
            for arg in func.args:
                arg.default = arg.deco = None

            func = cls.members["set_Item"]
            func.args[0].type = func.args[1].type = "int"
            func.args[2].type = "Range"
            for arg in func.args:
                arg.default = arg.deco = None

            cls.members["Value"] = cls.members.pop("Value2")
        elif cls.name == "DiagramNodeChildren":
            func = cls.members["AddNode"]
            func.args[0].type = "int"

        if "get_Item" in cls.members:
            func = cls.members["get_Item"]
            if len(func.args) == 1:
                arg = func.args[0]
                if arg.name in {"Index", "Level"} and arg.type == "object":
                    func.args[0].type = "int"
                elif arg.type == "int":
                    # does not require convert type
                    pass
                else:
                    assert True

        if "set_Item" in cls.members:
            func = cls.members["set_Item"]
            if len(func.args) == 2:
                arg = func.args[0]
                if arg.name in {"Index", "Level"} and arg.type == "object":
                    func.args[0].type = "int"
                    print(func)
                elif arg.type == "int":
                    # does not require convert type
                    pass
                else:
                    assert True
                    print(func)

        for key, value in cls.members.items():
            if isinstance(value, Function):
                func: Function = value
            elif isinstance(value, Property):
                prop: Property = value
                if prop.type in {"VBProject", "VBE"}:
                    prop.type = f"IVbe.{prop.type}"
                elif prop.type == "object" and prop.name == "ActiveSheet":
                    prop.type = ws_type

        assert "Get" not in cls.members
        assert "get" not in cls.members
        if "get_Item" in cls.members:
            cls.members["get"] = cls.members.pop("get_Item")

        assert "Set" not in cls.members
        assert "set" not in cls.members
        if "set_Item" in cls.members:
            cls.members["set"] = cls.members.pop("set_Item")

        if "PresetCamera" in cls.members:
            # ThreeDFormat
            cls.members["PresetCamera_"] = cls.members.pop("PresetCamera")

        members = {}
        for key, value in dict(cls.members).items():
            if key.startswith(("_", "Dummy")) or key == "this[]":
                continue

            members[camel_to_py_name(key, True)] = value

        cls.members = members

        for func in (node for node in cls.members.values() if isinstance(node, Function)):
            freeze_args = False
            required_args = []
            optinal_args = []
            default_args = []

            for arg in func.args:
                arg.name = camel_to_py_name(arg.name, False)

                assert arg is not None
                optinal = get_optinal_type(arg)
                if optinal is None:
                    assert not freeze_args
                    required_args.append(arg)
                elif optinal is OptinalType.HasDefine:
                    freeze_args = True
                    optinal_args.append(arg)
                elif optinal is OptinalType.HasDefault:
                    freeze_args = True
                    default_args.append(arg)

            func.py_args = tuple(required_args + optinal_args + default_args)

        return cls


class ClassFinder:
    def __init__(self, path: Path):
        self.path = path

    @lru_cache()
    def load(self, name) -> Class:
        if not self.has(name):
            raise FileNotFoundError(str(self.find(name)))

        return Parser(name, self).parse()

    @lru_cache()
    def find(self, name):
        return (self.path / f"{name}.cs")

    @lru_cache()
    def has(self, name):
        return self.find(name).exists()

    def read_code(self, name):
        assert self.has(name)
        if self.has(f"_{name}"):
            return self.find(f"_{name}").read_text(encoding="utf-8")

        return self.find(name).read_text(encoding="utf-8")


class Resolver:
    def __init__(self, finder: ClassFinder):
        self.finder = finder
        self.parsed = set()
        self.ignored = set()
        self.classes = set()
        self.enums = set()

    def load(self, cls_name):
        return self.finder.load(cls_name)

    def resolve(self, cls_name):
        if cls_name.islower() and cls_name in CS_DEFAULT_TYPES:
            return

        if not self.finder.has(cls_name):
            # print("!", cls_name)
            return None
        elif cls_name in self.ignored:
            return
        elif cls_name not in self.parsed:
            self.parsed.add(cls_name)
        else:
            return

        # print(cls_name)
        obj = self.finder.load(cls_name)
        if not ("class" in obj.header or "interface" in obj.header or "enum" in obj.header):
            self.parsed.remove(cls_name)
            self.ignored.add(cls_name)
            return

        if isinstance(obj, Class):
            self.classes.add(cls_name)
            for name, value in obj.members.items():
                if isinstance(value, Function):
                    func: Function = value
                    self.resolve(func.type)
                    for arg in func.args:
                        self.resolve(arg.type)
                elif isinstance(value, Property):
                    prop: Property = value
                    self.resolve(prop.type)
        elif isinstance(obj, Enum):
            self.enums.add(cls_name)
        else:
            raise Exception


class Buffer(io.StringIO):
    def __init__(self):
        super().__init__()
        self.indent = 0

    def write2(self, *args, sep=" ", end="\n"):
        # noinspection PyTypeChecker
        print(*args, sep=sep, end=end, file=self)

    def write_indent(self):
        # noinspection PyTypeChecker
        print("    " * self.indent, end="", file=self)

    # from pyi
    def write_start(self, *args, sep=" ", end="\n", indent_offset=0):
        self.indent += indent_offset
        self.write_indent()
        self.write2(*args, sep=sep, end=end)

    # from gen_cs
    def write_line(self, *args, sep=" ", end="\n"):
        self.write_indent()
        self.write2(*args, sep=sep, end=end)

    @contextmanager
    def gen_block(self):
        self.write_line("{")
        with self.with_indent():
            yield
        self.write_line("}")

    @contextmanager
    def with_indent(self):
        self.indent += 1
        try:
            yield
        finally:
            self.indent -= 1


class CsCompiler:
    def __init__(self, resolver: Resolver):
        self.resolver = resolver

    @property
    def finder(self):
        return self.resolver.finder

    def required_warp(self, type):
        return type in self.resolver.classes

    TYPES_CS2PY_MAP = {
        "object": "object",
        "string": "str",
        "int": "int",
        "bool": "bool",
        "void": None,
        "double": "float",
        "float": "float",
    }

    def type_cs_to_py(self, type):
        py_type = self.TYPES_CS2PY_MAP.get(type)
        if py_type is not None:
            return py_type

        if type in self.resolver.classes:
            return f"{type}Type"

        if type in self.resolver.enums:
            return str(type)

        return None

    def build_py_enum(self, enum: Enum):
        self.resolver.resolve(enum.name)

        buffer = Buffer()
        write_indent = buffer.write_indent
        write_line = buffer.write_line
        write = buffer.write2
        write_start = buffer.write_start

        # declear enum
        write()
        write_start(f"class {enum.name}(IntEnum):")
        with buffer.with_indent():
            assert enum.values

            for key, value in enum.values.items():
                write_line(key, "=", value)

            write()

        return buffer.getvalue()

    def build_pyi(self, cls: Class):
        self.resolver.resolve(cls.name)

        buffer = Buffer()
        write_indent = buffer.write_indent
        write_line = buffer.write_line
        write = buffer.write2
        write_start = buffer.write_start

        # declear class
        write_start(f"class {cls.name}:")
        with buffer.with_indent():
            if not cls.members:
                write_start("pass")
                write()

            get_func = cls.members.get("get")
            if get_func:
                write()
                write_indent()
                write(f"def __iter__(self", end="")
                for no, arg in enumerate(get_func.args):
                    write(", ", end="")

                    py_type = self.type_cs_to_py(arg.type)
                    if py_type:
                        write(f'{arg.name}: {py_type}', end="")
                    else:
                        write(f'{arg.name}', end="")

                write(")", end="")

                func_py_type = self.type_cs_to_py(get_func.type)
                if func_py_type:
                    write(f" -> Iterable[{func_py_type}]", end="")

                write(": pass")
                write()

            for alias_name, node in cls.members.items():
                if isinstance(node, Function):
                    func: Function = node

                    # function prototype
                    write()
                    write_indent()
                    write(f"def {alias_name}(self", end="")
                    for no, arg in enumerate(func.args):
                        write(", ", end="")

                        py_type = self.type_cs_to_py(arg.type)
                        if py_type:
                            write(f'{arg.name}: {py_type}', end="")
                        else:
                            write(f'{arg.name}', end="")

                        has_optional = get_optinal_type(arg) is not None
                        if has_optional:
                            write(" = None", end="")

                    write(")", end="")

                    func_py_type = self.type_cs_to_py(func.type)
                    if func_py_type:
                        write(f" -> {func_py_type}", end="")

                    write(": pass")
                    write()
                elif isinstance(node, Property):
                    prop: Property = node

                    write_indent()
                    py_type = self.type_cs_to_py(prop.type)
                    if py_type:
                        write(f'{alias_name}: {py_type}', end="")
                    else:
                        write(f'{alias_name}: None', end="")

                    write()

        write(f"{cls.name}Type = {cls.name}")
        write()
        write()

        return buffer.getvalue()

    def build_cs_class(self, cls: Class, obj_name):
        self.resolver.resolve(cls.name)

        buffer = Buffer()
        write_indent = buffer.write_indent
        write_line = buffer.write_line
        write = buffer.write2
        gen_block = buffer.gen_block

        def redirect_type(type):
            if type in self.resolver.classes:
                return type
            elif type in self.resolver.enums:
                return f"IExcel.{type}"

            return type

        write(f'// Auto-generated "{cls.name}.cs" by fastexcel.gen_cs v1.1')
        write()
        write("using System;")
        write("using System.Collections;")
        write("using System.Collections.Generic;")
        write("using System.Linq;")
        write("using System.Runtime.InteropServices;")
        write("using System.Runtime.CompilerServices;")
        write("using System.Text;")
        write("using System.Threading.Tasks;")
        write("using Microsoft.Office.Core;")
        write("using IExcel = Microsoft.Office.Interop.Excel;")
        write("using IVbe = Microsoft.Vbe.Interop;")
        write()

        namespace = "PyExcelLibrary"
        write(f"namespace {namespace}")
        with gen_block():
            # Function(type='object', name='get_Item', args=(Argument(type='object', name='RowIndex', io_type=None), Argument(type='object', name='ColumnIndex', io_type=None)))
            get_func: Function = cls.members.get("get")
            if get_func:
                if not (len(get_func.args) == 1 and
                        get_func.args[0].type in {"int", "object"} and
                        "count" in cls.members):
                    get_func = None

            cls_type = cls.name
            if cls_type in {"Workbook", "Application"}:
                cls_type = f"_{cls_type}"

            # declear class
            write_line(f"public partial class {cls.name} : ExcelObject<IExcel.{cls_type}>", end="")
            if get_func:
                write(f", IEnumerable<{get_func.type}>")
            else:
                write()

            with gen_block():
                if True:
                    # declear init
                    write_line(f"internal {cls.name}(IExcel.{cls_type} {obj_name}): base({obj_name})")
                    with gen_block():
                        pass

                if "name" in cls.members and cls.name != "Range":
                    write()
                    write_line(f"public override string ToString()")
                    with gen_block():
                        write_line(f'return ToStrTool.FixEncoding($"{cls.name}[{{ToStrTool.Quote(name)}}]");')

                if get_func:
                    write()
                    # function prototype
                    write_line(f"public IEnumerator<{get_func.type}> GetEnumerator()")
                    with gen_block():
                        # for loop
                        write_line("for (int i = 1; i <= count; i++)")
                        with gen_block():
                            write_line("yield return get(i);")

                    write()
                    write_line(f"IEnumerator IEnumerable.GetEnumerator()")
                    with gen_block():
                        write_line("return GetEnumerator();")

                for alias_name, node in cls.members.items():
                    if isinstance(node, Function):
                        func: Function = node

                        alias_name = escape_cs_func_name(alias_name)

                        func.type = redirect_type(func.type)

                        # function prototype
                        write()
                        write_line(f"public {func.type} {alias_name}(", end="")

                        for no, arg in enumerate(func.py_args):
                            if no:
                                write(", ", end="")

                            optional = get_optinal_type(arg)

                            if optional == OptinalType.HasDefine:
                                write("[Optional] ", end="")

                            arg_type = redirect_type(arg.type)
                            write(f'{arg_type} {arg.name}', end="")

                            if optional == OptinalType.HasDefault:
                                if arg_type != arg.type:
                                    write(f" = IExcel.{arg.default}", end="")
                                else:
                                    write(f" = {arg.default}", end="")

                        write(")")

                        with gen_block():
                            write_indent()

                            # return
                            if func.type != "void":
                                write("return", end=" ")

                                if self.required_warp(func.type):
                                    write(f"new {func.type}(", end="")

                            # call function
                            write(f"this.{obj_name}.{func.name}(", end="")
                            for no, arg in enumerate(func.args):
                                if no:
                                    write(", ", end="")

                                if arg.io_type is not None:
                                    write(arg.io_type, end=" ")

                                if self.required_warp(arg.type):
                                    write(f'{arg.name}.obj', end="")
                                else:
                                    write(f'{arg.name}', end="")

                            if self.required_warp(func.type):
                                write(")", end="")

                            # cleanup
                            write(");")
                    elif isinstance(node, Property):
                        prop: Property = node

                        alias_name = escape_cs_func_name(alias_name)

                        read_obj = f"this.{obj_name}.{prop.name}"
                        write_obj = f"this.{obj_name}.{prop.name} = value"

                        if self.required_warp(prop.type):
                            read_obj = f"new {prop.type}({read_obj})"
                            write_obj = f"{write_obj}.obj"

                        prop.type = redirect_type(prop.type)
                        prop_name = alias_name
                        if prop_name == cls.name:
                            prop_name = f"{prop_name}_"

                        write()
                        write_line(f"public {prop.type} {prop_name}", end="")
                        if prop.accessor == {"get"}:
                            write(f" => {read_obj};")
                        else:
                            write()
                            with gen_block():
                                if "get" in prop.accessor:
                                    write_indent()
                                    write(f"get => {read_obj};")
                                if "set" in prop.accessor:
                                    write_indent()
                                    write(f"set => {write_obj};")

        return buffer.getvalue()


def main():
    path = Path(r"Microsoft.Office.Interop.Excel")
    finder = ClassFinder(path)
    resolver = Resolver(finder)
    resolver.resolve("Application")

    compiler = CsCompiler(resolver)

    with (Path(__file__).parent / "../execlib/enums.py").open("w", encoding="utf-8") as out:
        out.write("from enum import IntEnum\n")
        out.write("\n")
        out.write("__all__ = " + repr([clsname for clsname in sorted(resolver.enums)]) + "\n")
        out.write("\n")
        for clsname in resolver.enums:
            obj = finder.load(clsname)
            assert isinstance(obj, Enum)
            out.write(compiler.build_py_enum(obj))

    with (Path(__file__).parent / "../execlib/PyExcelLibrary.pyi").open("w", encoding="utf-8") as out:
        out.write("from typing import *\n")
        out.write("from .enums import *\n")
        out.write("\n")
        out.write("__all__ = " + repr([clsname for clsname in sorted(resolver.classes)]) + "\n")
        out.write("\n")
        for clsname in resolver.classes:
            cls = finder.load(clsname)
            assert isinstance(cls, Class)
            out.write(compiler.build_pyi(cls))

    OUTPUT_PATH = Path("PyExcelLibrary/excel")
    for file in OUTPUT_PATH.glob("*.cs"):
        file.unlink()

    for clsname in resolver.classes:
        obj = finder.load(clsname)
        assert isinstance(obj, Class)
        with (OUTPUT_PATH / (clsname + ".cs")).open("w", encoding="utf-8") as out:
            out.write(compiler.build_cs_class(obj, "obj"))


# TODO: add support enum?

if __name__ == '__main__':
    main()
