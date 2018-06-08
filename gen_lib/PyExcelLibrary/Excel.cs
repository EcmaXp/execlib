using Python.Runtime;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using static Python.Runtime.Runtime;
using IExcel = Microsoft.Office.Interop.Excel;
using IVbe = Microsoft.Vbe.Interop;

namespace PyExcelLibrary
{
    internal class ToStrTool
    {
        internal static bool HasEncodingProblem;

        static ToStrTool()
        {
            using (Py.GIL())
            {
                var scope = Py.CreateScope();
                scope.Set("instance", new ToStrToolTest());
                HasEncodingProblem = scope.Eval("len(str(instance)) != 10").IsTrue();
            }

            ;
        }

        internal static string FixEncoding(string str)
        {
            if (HasEncodingProblem)
            {
                return str.PadRight(Encoding.UTF8.GetByteCount(str), '\0');
            }
            else
            {
                return str;
            }
        }

        internal static string Quote(string str)
        {
            return "\"" + str.Replace("\"", "\\\"") + "\"";
        }

        internal static string Quote(object str)
        {
            return Quote(str as string);
        }

        public class ToStrToolTest
        {
            public override string ToString()
            {
                return "안녕PADDINGX"; // len=10
            }
        }
    }

    public class ExcelType
    {
        public static readonly System.Object[] @Array1D = new System.Object[0];
        public static readonly System.Object[,] @Array2D = new System.Object[0, 0];
    }

    public class ExcelObject<T> : Object
    {
        internal T obj;

        public T _object_ => obj;

        internal ExcelObject(T obj)
        {
            if (obj == null)
            {
                throw new ArgumentNullException();
            }

            this.obj = obj;
        }
    }

    public static class PythonExtension
    {
        internal static PyObject PyNone;
        internal static PyObject PyTrue;
        internal static PyObject PyBoolType;
        internal static PyObject PySliceType;

        static PythonExtension()
        {
            using (Py.GIL())
            {
                PyObject module = Py.Import("builtins");
                PyNone = module.GetAttr("None");
                PyTrue = module.GetAttr("True");
                PyBoolType = PyTrue.GetPythonType();
                PySliceType = module.GetAttr("slice");
            }
        }

        public static bool IsNone(this PyObject obj)
        {
            return obj.IsInstance(PyNone.GetPythonType());
        }

        public static bool IsInt(this PyObject obj)
        {
            return PyInt.IsIntType(obj);
        }

        public static bool IsString(this PyObject obj)
        {
            return PyString.IsStringType(obj);
        }

        public static bool IsSlice(this PyObject obj)
        {
            return obj.IsInstance(PySliceType);
        }
    }

    public partial class Range
    {
        public override string ToString()
        {
            var address = this.get_address(row_absolute: false, column_absolute: false, external: true);
            return ToStrTool.FixEncoding($"Range[{ToStrTool.Quote(address)})]");
        }

        public Range get_item0_key(PyObject key)
        {
            using (Py.GIL())
            {
                if (key.IsString())
                {
                    return get_range(key.As<string>());
                }

                if (key.IsSlice())
                {
                    var range = new PyRangeAccess(key);
                    if (range.Start == null && range.Stop == null)
                    {
                        return this;
                    }

                    if ((range.Start?.IsString() ?? false) &&
                        (range.Stop?.IsString() ?? false))
                    {
                        return get_range($"{range.Start.As<string>()}{range.Stop.As<string>()}");
                    }
                }

                throw new Exception("invaild type");
            }
        }

        internal Range get_range(IntRange rowRange, IntRange columnRange)
        {
            if (rowRange.Start >= 0 && columnRange.Start >= 0)
            {
                return get_range(
                    $"{GetColumnString(columnRange.Start)}{rowRange.Start}",
                    $"{GetColumnString(columnRange.Stop)}{rowRange.Stop}"
                );
            }

            return get_offset(rowRange.Start, columnRange.Start)
                .get_resize(rowRange.Length, columnRange.Length);

        }

        internal IntRange _CalcCellRange(PyObject key, Range seq)
        {
            if (key.IsNone())
            {
                return new IntRange(1, seq.count);
            }
            
            if (key.IsInt())
            {
                return new IntRange(
                    key.As<int>() + 1
                );
            }

            if (key.IsSlice())
            {
                var range = new PyRangeAccess(key);
                if (range.Step != null)
                    throw new Exception();

                return new IntRange(
                    range.Start?.As<int>() + 1 ?? 1,
                    range.Stop?.As<int>() ?? rows.count
                );
            }

            throw new Exception($"key = {key}");
        }

        public Range get_item0(PyObject kRow, PyObject kColumn)
        {
            using (Py.GIL())
            {
                return get_range(
                    _CalcCellRange(kRow, rows),
                    _CalcCellRange(kColumn, columns)
                );
            }
        }

        internal static string GetColumnString(int n)
        {
            n -= 1;
            if (0 <= n && n < 26)
            {
                char[] r =
                {
                    (char) ('A' + n)
                };
                return new String(r);
            }

            if (26 <= n && n < 702)
            {
                char[] r =
                {
                    (char) ('A' + (n / 26) - 1),
                    (char) ('A' + (n % 26)),
                };
                return new String(r);
            }

            if (702 <= n && n < 16384)
            {
                n -= 702;
                char[] r =
                {
                    (char) ('A' + (n / 676)),
                    (char) ('A' + (n % 676 / 26)),
                    (char) ('A' + (n % 26)),
                };
                return new String(r);
            }

            throw new Exception("too large n");
        }
    }

    internal struct IntRange
    {
        public readonly int Start;
        public readonly int Stop;

        public IntRange(int num)
        {
            this.Start = this.Stop = num;
        }

        public IntRange(int start, int stop)
        {
            this.Start = start;
            this.Stop = stop;
        }

        public int Length => Stop - Start;
    }
    
    internal class PyRangeAccess
    {
        // ReSharper disable once InconsistentNaming
        internal PyObject obj { get; }

        public PyRangeAccess(PyObject obj)
        {
            this.obj = obj;
        }

        public PyObject Start
        {
            get
            {
                PyObject ob = obj.GetAttr("start");
                return ob.IsNone() ? null : ob;
            }
        }

        public PyObject Stop
        {
            get
            {
                PyObject ob = obj.GetAttr("stop");
                return ob.IsNone() ? null : ob;
            }
        }

        public PyObject Step
        {
            get
            {
                PyObject ob = obj.GetAttr("step");
                return ob.IsNone() ? null : ob;
            }
        }
    }
}