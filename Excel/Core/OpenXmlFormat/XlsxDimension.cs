using System;

namespace ExcelDataReader.Portable.Core.OpenXmlFormat
{
    internal class XlsxDimension
    {
        public XlsxDimension(string value, string overrideLastColumn = "XEN")
        {
            ParseDimensions(value, overrideLastColumn);
        }

        public XlsxDimension(int rows, int cols)
        {
            this.FirstRow = 1;
            this.LastRow = rows;
            this.FirstCol = 1;
            this.LastCol = cols;
        }

        private int _FirstRow;

        public int FirstRow
        {
            get { return _FirstRow; }
            set { _FirstRow = value; }
        }

        private int _LastRow;

        public int LastRow
        {
            get { return _LastRow; }
            set { _LastRow = value; }
        }

        private int _FirstCol;

        public int FirstCol
        {
            get { return _FirstCol; }
            set { _FirstCol = value; }
        }

        private int _LastCol;

        public int LastCol
        {
            get { return _LastCol; }
            set { _LastCol = value; }
        }

        public void ParseDimensions(string value, string overrideLastColumn = "XEN")
        {
            string[] parts = value.Split(':');

            int col;
            int row;

            XlsxDim(parts[0], out col, out row);
            FirstCol = col;
            FirstRow = row;

            if (parts.Length == 1)
            {
                LastCol = FirstCol;
                LastRow = FirstRow;
            }
            else
            {
                if (!string.IsNullOrEmpty(overrideLastColumn))
                {
                    parts[1] = ReplaceColumn(parts[1], overrideLastColumn);
                }

                XlsxDim(parts[1], out col, out row);
                LastCol = col;
                LastRow = row;
            }
        }

        /// <summary>
        /// Replace the end column number.
        /// For example, incoming value may be: XEN20, but we are overriding the column with W
        /// so the result would be: W20.
        /// </summary>
        /// <param name="value">
        /// The value.
        /// </param>
        /// <param name="overrideLastColumn">
        /// The override last column.
        /// </param>
        /// <returns>
        /// The <see cref="string"/>.
        /// </returns>
        private string ReplaceColumn(string value, string overrideLastColumn = "XEN")
        {
            string result = overrideLastColumn;
            int index = 0;
            while (index < value.Length)
            {
                if (char.IsDigit(value[index]))
                {
                    result += value[index];
                }

                index++;
            }

            return result;
        }

        /// <summary>
        /// Logic for the Excel dimensions. Ex: A15
        /// </summary>
        /// <param name="value">The value.</param>
        /// <param name="val1">out val1.</param>
        /// <param name="val2">out val2.</param>
        public static void XlsxDim(string value, out int val1, out int val2)
        {//INFO: Check for a simple Solution
            int index = 0;
            val1 = 0;
            int[] arr = new int[value.Length - 1];

            while (index < value.Length)
            {
                if (char.IsDigit(value[index])) break;
                arr[index] = value[index] - 'A' + 1;
                index++;
            }

            for (int i = 0; i < index; i++)
            {
                val1 += (int)(arr[i] * Math.Pow(26, index - i - 1));
            }

            val2 = int.Parse(value.Substring(index));
        }
    }
}
