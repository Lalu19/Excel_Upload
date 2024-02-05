using ClosedXML.Excel;
using Projectt.Models.Database;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;

namespace Projectt.Controllers
{
	public class HomeController : Controller
	{

		private ImportExcelEntities db =new ImportExcelEntities();
		public ActionResult Index()
		{
			return View();
		}

		public ActionResult UploadFile(HttpPostedFileBase myExcelData)
		{
			if (myExcelData != null && myExcelData.ContentLength > 0)
			{
				string fileName = $"{Guid.NewGuid()}.xlsx";
				string filePath = Path.Combine("D:\\Lalu\\Assignment\\Projectt\\Projectt\\Storage\\Upload\\", fileName);

				//filePath = filePath + fileName + ".xlsx";
				myExcelData.SaveAs(filePath);

				// download link
				string downloadLink = Url.Action("DownloadFile", "Home", new { fileName = fileName });

				XLWorkbook xlworkbook = new XLWorkbook(filePath);
				int row = 2;
				while (xlworkbook.Worksheets.Worksheet(1).Cell(row, 1).GetString() != "")
				{
					string slno = xlworkbook.Worksheets.Worksheet(1).Cell(row, 1).GetString();
					string name = xlworkbook.Worksheets.Worksheet(1).Cell(row, 2).GetString();
					string email = xlworkbook.Worksheets.Worksheet(1).Cell(row, 3).GetString();
					string phoneno = xlworkbook.Worksheets.Worksheet(1).Cell(row, 4).GetString();
					string address = xlworkbook.Worksheets.Worksheet(1).Cell(row, 5).GetString();
					string status = xlworkbook.Worksheets.Worksheet(1).Cell(row, 6).GetString();

					// Validate email format using regular expression
					if (!IsValidEmail(email))
					{
						return Json(new { success = false, message = "Invalid email format in the Excel file" }, JsonRequestBehavior.AllowGet);
					}

					Uploaddata up = new Uploaddata();
					up.Id = Guid.NewGuid();
					up.Slno = slno;
					up.Name = name;
					up.Email = email;
					up.PhoneNo = phoneno;
					up.Address = address;
					up.Status = status;

					db.Uploaddata.Add(up);
					db.SaveChanges();

					row++;
				}

				return Json(new { success = true, message = "Success", downloadLink }, JsonRequestBehavior.AllowGet);
			}
		
			else
			{
				return Json(new { success = false, message = "Please upload an excel file" }, JsonRequestBehavior.AllowGet);
			}
			return Json(new { success = true, message = "Success" }, JsonRequestBehavior.AllowGet);



		}

		private bool IsValidEmail(string email)
		{
			string pattern = @"^[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$";
			Regex regex = new Regex(pattern);
			return regex.IsMatch(email);
		}

		public ActionResult DownloadFile(string fileName)
		{
			string filePath = Path.Combine("D:\\Lalu\\Assignment\\Projectt\\Projectt\\Storage\\Upload\\", fileName);

			if (System.IO.File.Exists(filePath))
			{
				byte[] fileBytes = System.IO.File.ReadAllBytes(filePath);
				return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, fileName);
			}
			else
			{
				return HttpNotFound(); 
			}
		}

		public ActionResult About()
		{
			ViewBag.Message = "Your application description page.";

			return View();
		}

		public ActionResult Contact()
		{
			ViewBag.Message = "Your contact page.";

			return View();
		}
	}
}