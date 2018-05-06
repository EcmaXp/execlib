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


def parse_body(lines):
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

CLASS_TREE = {
    "Application": "Application",
    "Workbooks": "Application",
    "Workbook": "Workbook",
    "Sheets": "Workbook",
    "Worksheet": "Sheets",
    "Range": "Worksheet",
}


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

        code = self.read_code(name)
        namespace, header, lines = parse_headers(code)
        members = {node.name: node for node in parse_body(lines)}
        cls = Class(name, members, namespace, header)
        cls = patch_cls(cls)
        return cls

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
        cls = self.finder.load(cls_name)
        if not ("class" in cls.header or "interface" in cls.header):
            self.parsed.remove(cls_name)
            self.ignored.add(cls_name)
            return

        for name, value in cls.members.items():
            if isinstance(value, Function):
                func: Function = value
                self.resolve(func.type)
                for arg in func.args:
                    self.resolve(arg.type)
            elif isinstance(value, Property):
                prop: Property = value
                self.resolve(prop.type)


class CsCompiler:
    def __init__(self, resolver: Resolver):
        self.resolver = resolver

    @property
    def finder(self):
        return self.resolver.finder

    def redirect_input_type(self, type):
        # TODO: alloweed type?

        if type in self.resolver.parsed:
            return f"Excel.{type}"

        return type

    def required_warp(self, type):
        return type in self.resolver.parsed

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
        return self.TYPES_CS2PY_MAP.get(type, f"{type}Type" if type in self.resolver.parsed else None)

    def build_pyi(self, cls: Class):
        self.resolver.resolve(cls.name)
        fp = io.StringIO()  # type: Union[io.StringIO, IO[str]]

        indent = 0

        def write_indent():
            print("    " * indent, end="", file=fp)

        def write(*args, sep=" ", end="\n"):
            print(*args, sep=sep, end=end, file=fp)

        def write_start(*args, sep=" ", end="\n", indent_offset=0):
            nonlocal indent
            indent += indent_offset
            write_indent()
            write(*args, sep=sep, end=end)

        # declear class
        write_start(f"class {cls.name}:")
        indent += 1

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

        indent -= 1
        write(f"{cls.name}Type = {cls.name}")
        write()
        write()

        return fp.getvalue()

    def build_cs_class(self, cls: Class, obj_name):
        self.resolver.resolve(cls.name)
        required_warp = self.required_warp
        fp = io.StringIO()  # type: Union[io.StringIO, IO[str]]

        indent = 0

        def write_indent():
            print("    " * indent, end="", file=fp)

        def write(*args, sep=" ", end="\n"):
            print(*args, sep=sep, end=end, file=fp)

        def write_line(*args, sep=" ", end="\n"):
            write_indent()
            write(*args, sep=sep, end=end)

        @contextmanager
        def gen_block():
            nonlocal indent
            write_line("{")
            indent += 1
            try:
                yield
            finally:
                indent -= 1

            write_line("}")

        def redirect_enum_type(type):
            return type if type not in self.resolver.ignored else f"IExcel.{type}"

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

                        func.type = redirect_enum_type(func.type)

                        # function prototype
                        write()
                        write_line(f"public {func.type} {alias_name}(", end="")

                        for no, arg in enumerate(func.py_args):
                            if no:
                                write(", ", end="")

                            optional = get_optinal_type(arg)

                            if optional == OptinalType.HasDefine:
                                write("[Optional] ", end="")

                            arg_type = redirect_enum_type(arg.type)
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

                        prop.type = redirect_enum_type(prop.type)
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

        return fp.getvalue()


def main():
    path = Path(r"Microsoft.Office.Interop.Excel")
    finder = ClassFinder(path)
    resolver = Resolver(finder)
    resolver.resolve("Application")

    compiler = CsCompiler(resolver)

    with (Path(__file__).parent / "../execlib/PyExcelLibrary.pyi").open("w", encoding="utf-8") as out:
        out.write("from typing import *\n")
        out.write("\n")
        out.write("__all__ = " + repr([clsname for clsname in sorted(resolver.parsed)]) + "\n")
        out.write("\n")
        for clsname in resolver.parsed:
            out.write(compiler.build_pyi(finder.load(clsname)))

    OUTPUT_PATH = Path("PyExcelLibrary/excel")
    for file in OUTPUT_PATH.glob("*.cs"):
        file.unlink()

    for clsname in resolver.parsed:
        with (OUTPUT_PATH / (clsname + ".cs")).open("w", encoding="utf-8") as out:
            out.write(compiler.build_cs_class(finder.load(clsname), "obj"))


if __name__ == '__main__':
    main()
