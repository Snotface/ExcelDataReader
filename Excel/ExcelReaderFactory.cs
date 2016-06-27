using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace Excel
{
	/// <summary>
	/// The ExcelReader Factory
	/// </summary>
	public static class ExcelReaderFactory
	{

		/// <summary>
		/// Creates an instance of <see cref="ExcelBinaryReader"/>
		/// </summary>
		/// <param name="fileStream">The file stream.</param>
		/// <returns></returns>
		public static IExcelDataReader CreateBinaryReader(Stream fileStream, string overrideLastColumn = "")
		{
			IExcelDataReader reader = new ExcelBinaryReader();
			reader.Initialize(fileStream, overrideLastColumn);

			return reader;
		}

		/// <summary>
		/// Creates an instance of <see cref="ExcelBinaryReader"/>
		/// </summary>
		/// <param name="fileStream">The file stream.</param>
		/// <returns></returns>
		public static IExcelDataReader CreateBinaryReader(Stream fileStream, ReadOption option, string overrideLastColumn = "")
		{
			IExcelDataReader reader = new ExcelBinaryReader(option);
			reader.Initialize(fileStream, overrideLastColumn);

			return reader;
		}

		/// <summary>
		/// Creates an instance of <see cref="ExcelBinaryReader"/>
		/// </summary>
		/// <param name="fileStream">The file stream.</param>
		/// <returns></returns>
		public static IExcelDataReader CreateBinaryReader(Stream fileStream, bool convertOADate, string overrideLastColumn = "")
		{
			IExcelDataReader reader = CreateBinaryReader(fileStream, overrideLastColumn);
			((ExcelBinaryReader) reader).ConvertOaDate = convertOADate;

			return reader;
		}

		/// <summary>
		/// Creates an instance of <see cref="ExcelBinaryReader"/>
		/// </summary>
		/// <param name="fileStream">The file stream.</param>
		/// <returns></returns>
		public static IExcelDataReader CreateBinaryReader(Stream fileStream, bool convertOADate, ReadOption readOption, string overrideLastColumn = "")
		{
			IExcelDataReader reader = CreateBinaryReader(fileStream, readOption, overrideLastColumn);
			((ExcelBinaryReader)reader).ConvertOaDate = convertOADate;

			return reader;
		}

		/// <summary>
		/// Creates an instance of <see cref="ExcelOpenXmlReader"/>
		/// </summary>
		/// <param name="fileStream">The file stream.</param>
		/// <returns></returns>
		public static IExcelDataReader CreateOpenXmlReader(Stream fileStream, string overrideLastColumn = "")
		{
			IExcelDataReader reader = new ExcelOpenXmlReader();
			reader.Initialize(fileStream, overrideLastColumn);

			return reader;
		}
	}
}
