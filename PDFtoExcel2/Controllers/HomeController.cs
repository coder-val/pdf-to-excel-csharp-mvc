using PDFtoExcel2.utils;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Hosting;
using System.Web.Mvc;
using System.Web.UI.WebControls;
using iText.Kernel.Pdf;

namespace PDFtoExcel2.Controllers 
{
	public class HomeController : Controller
	{

		public ActionResult Index()
		{
            return View();
		}

		[HttpPost]
		public ActionResult UploadFile(HttpPostedFileBase fileUpload)
		{
			if (fileUpload != null && fileUpload.ContentLength > 0)
			{
				try
				{
					string dt = DateTime.Now.ToString("MMddyHHmmss");
					char[] s = { '.' };
					string fileName = Path.GetFileName(fileUpload.FileName);
					string[] sFileName = fileName.Split(s);
					string newFileName = $"{sFileName[0]}_{dt}.{sFileName[1]}";
					string filePath = Path.Combine(Server.MapPath("~/media/uploads"), newFileName);
					fileUpload.SaveAs(filePath);

					var util = new Utilities();
					util.PDFtoText(newFileName, new PdfDocument(new PdfReader(filePath)));
                    ViewBag.Color = "success";
                    ViewBag.Message = "SUCCESS";

					string[] snewFileName = newFileName.Split(s);
                    ViewBag.excel = $"{snewFileName[0]}.xlsx";
                    ViewBag.pdf = $"{snewFileName[0]}.pdf";
                    return View("Index");
                    //return RedirectToAction("Index", "Home", new { fileName = newFileName});
                }
				catch
				{
                    ViewBag.Color = "danger";
                    ViewBag.Message = "Something went wrong, please check your file.";
                    return View("Index");
                }
			}
			else
			{
				ViewBag.Color = "warning";
				ViewBag.Message = "Please select a file to upload";
                return View("Index");
            }
		}

		public ActionResult DownloadExcelFile(string fileName)
		{
            string filePath = Path.Combine(Server.MapPath("~/media/downloads/excel"), fileName);
			byte[] fileBytes = System.IO.File.ReadAllBytes(filePath);
			string file = fileName;
			return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, file);
		}
        public ActionResult DownloadPDFFile(string fileName)
        {
            string filePath = Path.Combine(Server.MapPath("~/media/downloads/pdf"), fileName);
            byte[] fileBytes = System.IO.File.ReadAllBytes(filePath);
            string file = fileName;
            return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Pdf, file);
        }
	}
}