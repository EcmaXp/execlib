using Python.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PyExcelLibrary
{
    internal class ToStrTool
    {
        internal static bool HasEncodingProblem;

        static ToStrTool() {
            using (Py.GIL())
            {
                var scope = Py.CreateScope();
                scope.Set("instance", new ToStrToolTest());
                HasEncodingProblem = scope.Eval("len(str(instance)) != 10").IsTrue();
            };
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

    public class ExcelObject<T>: Object
    {
        internal T obj;

        internal ExcelObject(T obj)
        {
            this.obj = obj;
        }
    }

    public partial class Range
    {
        public override string ToString()
        {
            return ToStrTool.FixEncoding($"Range[{ToStrTool.Quote(this.get_address(row_absolute: false, column_absolute: false, external: true))}]");
        }
    }
}
