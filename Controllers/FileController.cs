using Excel2JsonApi.Extensions;
using ExcelDataReader;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace Excel2JsonApi.Controllers
{
    [Route("api/[controller]/[action]")]
    public class FileController : ControllerBase
    {
        private readonly IConfiguration _configuration;

        public FileController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpGet]
        public ActionResult Resub() =>
            Ok(new DirectoryInfo(_configuration["ResubPath"])
                .EnsureCreatedAndEnumerateFiles()
                .Select(file => new { file.Name, file.FullName }));

        [HttpGet]
        public ActionResult Reprocess() =>
            Ok(new DirectoryInfo(_configuration["ReprocessPath"])
                .EnsureCreatedAndEnumerateFiles()
                .Select(file => new { file.Name, file.FullName }));

        [HttpGet]
        public ActionResult Process() =>
            Ok(new DirectoryInfo(_configuration["ProcessPath"])
                .EnsureCreatedAndEnumerateFiles()
                .Select(file => new { file.Name, file.FullName }));

        [HttpGet]
        public ActionResult ConvertToJson([FromQuery] string filePath)
        {
            FileInfo file = new FileInfo(filePath);

            if (!file.Exists)
            {
                return NotFound();
            }

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var config = new ExcelReaderConfiguration
            {
                FallbackEncoding = Encoding.GetEncoding(1252)
            };

            using (var fileStream = file.OpenRead())
            using (var reader = ExcelReaderFactory.CreateBinaryReader(fileStream, config))
            {
                var dsConfig = new ExcelDataSetConfiguration
                {
                    UseColumnDataType = true,
                    ConfigureDataTable = tableReader => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true
                    }
                };

                var ds = reader.AsDataSet(dsConfig);

                List<object> tables = new List<object>();

                foreach (DataTable dataTable in ds.Tables)
                {
                    List<Dictionary<string, object>> rows = new List<Dictionary<string, object>>();
                    List<string> cols = new List<string>();

                    foreach (DataColumn col in dataTable.Columns)
                    {
                        cols.Add(col.ColumnName);
                    }

                    foreach (DataRow dataRow in dataTable.Rows)
                    {
                        Dictionary<string, object> row = new Dictionary<string, object>();

                        foreach (DataColumn dataCol in dataTable.Columns)
                        {
                            row.Add(dataCol.ColumnName, Convert.ChangeType(dataRow[dataCol], dataCol.DataType));
                        }

                        rows.Add(row);
                    }

                    tables.Add(new { dataTable.TableName, Data = new { Columns = cols, Rows = rows } });
                }

                return Ok(tables);
            }
        }
    }
}
