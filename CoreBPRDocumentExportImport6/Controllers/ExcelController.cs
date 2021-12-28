using System.Data;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Dapper;
using Npgsql;
using Syncfusion.XlsIO;
using CoreBPRDocumentExportImport6.Models;

namespace CoreBPRDocumentExportImport6.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        private readonly ApplicationDbContext _context;
        private IConfiguration Configuration { get; }
        private readonly ILogger<ExcelController> _logger;

        public ExcelController(ApplicationDbContext context, IConfiguration configuration, ILogger<ExcelController> logger)
        {
            _context = context;
            Configuration = configuration;
            _logger = logger;
        }

        [HttpGet("lapkeu/{template}")]
        public async Task<ActionResult<string>> GetLapkeu(string template)
        {
            _logger.LogInformation("GetLapkeu started at: {time}", DateTimeOffset.UtcNow);

            try
            {
                string filePath = await GenerateLaporanKeuangan(template);
                return Ok(filePath);
            }
            catch (Exception e)
            {
                return BadRequest(e.Message);
            }
        }

        // Generate Laporan Keuangan
        private async Task<string> GenerateLaporanKeuangan(string template)
        {
            _logger.LogInformation("GenerateLaporanKeuangan started at: {time}", DateTimeOffset.UtcNow);

            string templateFolder = Configuration["Folder:TemplateFolder"];
            string uploadFolder = Configuration["Folder:UploadFolder"];
            string templateFileName = _context.DcxTemplateMasters.Where(x => x.TemplateId == template).Select(x => x.TemplateFilename).FirstOrDefault();
            string templateFile = Path.Combine(templateFolder, templateFileName);
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(templateFile);
            string fileExtension = Path.GetExtension(templateFile);
            DateTime today = DateTime.Now;
            string strToday = today.ToString("yyyy-MM-dd");
            string createdFileName = fileNameWithoutExtension + "-" + strToday + fileExtension;
            string createdFile = Path.Combine(uploadFolder, createdFileName);

            _logger.LogInformation("templateFolder: {0}", templateFolder);
            _logger.LogInformation("uploadFolder: {0}", uploadFolder);
            _logger.LogInformation("templateFileName: {0}", templateFileName);
            _logger.LogInformation("templateFile: {0}", templateFile);
            _logger.LogInformation("fileNameWithoutExtension: {0}", fileNameWithoutExtension);
            _logger.LogInformation("fileExtension: {0}", fileExtension);
            _logger.LogInformation("strToday: {0}", strToday);
            _logger.LogInformation("createdFileName: {0}", createdFileName);
            _logger.LogInformation("createdFile: {0}", createdFile);

            try
            {
                if (System.IO.File.Exists(createdFile))
                {
                    int i = 1;
                    string newFileName = fileNameWithoutExtension + "-" + strToday + "(" + i.ToString() + ")" + fileExtension;
                    string newFilePath = Path.Combine(uploadFolder, newFileName);
                    while (System.IO.File.Exists(newFilePath))
                    {
                        i++;
                        newFileName = fileNameWithoutExtension + "-" + strToday + "(" + i.ToString() + ")" + fileExtension;
                        newFilePath = Path.Combine(uploadFolder, newFileName);
                    }
                    createdFileName = newFileName;
                    createdFile = newFilePath;

                    _logger.LogInformation("File Exists");
                    _logger.LogInformation("createdFileName: {0}", createdFileName);
                    _logger.LogInformation("createdFile: {0}", createdFile);
                }

                List<DcxTemplateMaster> listDcxTemplateMaster;
                List<DcxTemplateDetail> listDcxTemplateDetail;
                string sheetId, sqlSelect, sqlFrom, sqlWhere, sqlQuery, pos, saldo;
                int sequenceId;
                DataTable tableResult;

                using (ExcelEngine excelEngine = new ExcelEngine())
                {
                    IApplication application = excelEngine.Excel;

                    FileStream templateFileStream = new FileStream(templateFile, FileMode.OpenOrCreate, FileAccess.Read);

                    //Loads or open an existing workbook through Open method of IWorkbooks
                    IWorkbook workbook = excelEngine.Excel.Workbooks.Open(templateFileStream);

                    // Loop for DcxTemplateMaster
                    listDcxTemplateMaster = _context.DcxTemplateMasters.Where(x => x.TemplateId == template).OrderBy(x => x.Id).ToList();

                    foreach (DcxTemplateMaster dcxTemplateMaster in listDcxTemplateMaster)
                    {
                        sheetId = dcxTemplateMaster.SheetId;
                        sequenceId = dcxTemplateMaster.SequenceId;

                        //Access a worksheet from workbook
                        IWorksheet worksheet = workbook.Worksheets[sheetId];

                        // Loop for DcxTemplateDetail
                        listDcxTemplateDetail = _context.DcxTemplateDetails.Where(x => x.TemplateId == template && x.SheetId == sheetId && x.SequenceId == sequenceId).OrderBy(x => x.Id).ToList();

                        sqlSelect = "";
                        sqlFrom = "";
                        sqlWhere = "";
                        sqlQuery = "";

                        foreach (DcxTemplateDetail dcxTemplateDetail in listDcxTemplateDetail)
                        {
                            if ((dcxTemplateDetail.SqlSelect != null) && (dcxTemplateDetail.SqlSelect.Trim() != String.Empty))
                            {
                                sqlSelect = sqlSelect + dcxTemplateDetail.SqlSelect;
                            }

                            if ((dcxTemplateDetail.SqlFrom != null) && (dcxTemplateDetail.SqlFrom.Trim() != String.Empty))
                            {
                                sqlFrom = sqlFrom + dcxTemplateDetail.SqlFrom;
                            }

                            if ((dcxTemplateDetail.SqlWhere != null) && (dcxTemplateDetail.SqlWhere.Trim() != String.Empty))
                            {
                                sqlWhere = sqlWhere + dcxTemplateDetail.SqlWhere;
                            }
                        }

                        sqlQuery = sqlSelect + sqlFrom + sqlWhere;

                        _logger.LogInformation("sqlQuery: {0}", sqlQuery);

                        // Dapper Connection
                        using (var dconn = new NpgsqlConnection(Configuration.GetConnectionString("DefaultConnection")))
                        {
                            var queryResult = dconn.Query(sqlQuery).ToList();

                            var jsonResult = JsonConvert.SerializeObject(queryResult);

                            _logger.LogInformation("jsonResult: {0}", jsonResult);

                            tableResult = (DataTable)JsonConvert.DeserializeObject(jsonResult, typeof(DataTable));

                            for (int i = 0; i < tableResult.Rows.Count; i++)
                            {
                                pos = tableResult.Rows[i][0].ToString();
                                saldo = tableResult.Rows[i][1].ToString();

                                worksheet.Range[pos].Text = saldo;
                            }
                        }
                    }

                    //Saving the workbook to disk in XLSX format
                    FileStream result = new FileStream(createdFile, FileMode.Create, FileAccess.Write);
                    workbook.SaveAs(result);

                    // Clean Process
                    workbook.Close();
                    templateFileStream.Close();
                    result.Close();

                    return createdFileName;
                }
            }
            catch (Exception e)
            {
                _logger.LogError("catch: {0}", e.Message);

                if (e.Message.Contains(templateFile))
                {
                    throw new Exception(e.Message.Replace(templateFile, templateFileName));
                }
                if (e.Message.Contains(createdFile))
                {
                    throw new Exception(e.Message.Replace(createdFile, createdFileName));
                }
                throw new Exception(e.Message);
            }
        }
    }
}
