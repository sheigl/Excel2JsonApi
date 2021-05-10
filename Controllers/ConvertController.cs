using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using ExcelDataReader;
using Newtonsoft.Json;
using System.Text;
using System.Collections.Generic;
using System.Data;
using System;
using System.IO;

namespace Excel2json.Controllers
{
    [Route("api/[controller]")]
    public class ConvertController : ControllerBase
    {
        public class UploadModel
        {
            public string Payload { get; set; }
        }


        [HttpPost]
        public ActionResult Post([FromBody] UploadModel payload)
        {
            if (String.IsNullOrEmpty(payload.Payload))
                return BadRequest("No file uploaded");            

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            var config = new ExcelReaderConfiguration 
            {
                 FallbackEncoding = Encoding.GetEncoding(1252)
            };

            using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(payload.Payload)))
            using (var reader = ExcelReaderFactory.CreateBinaryReader(ms, config))
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

                    tables.Add(new { dataTable.TableName, Data = new { Columns = cols, Rows = rows }});
                }

                return Ok(tables);
            }
        }
    }
}