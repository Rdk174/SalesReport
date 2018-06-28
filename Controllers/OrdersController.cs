using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using SalesReport.Models;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Globalization;
using System.IO;
using System.Diagnostics;
using System.Net.Mail;

namespace SalesReport.Controllers
{
    public class OrdersController : Controller
    {
        private NorthwindEntities db = new NorthwindEntities();
        // GET: Orders
        //public ActionResult Index()
        //{
        //    var orderDetail = (from od in db.OrderDetail
        //                       join o in db.Order on od.OrderID equals o.ID
        //                       join p in db.Product on od.ProductID equals p.ID
        //                       select new Orders
        //                       {
        //                           OrderID = od.OrderID,
        //                           OrderDate = o.OrderDate,
        //                           ProductId = od.OrderID,
        //                           ProductName = p.Name,
        //                           Quntity = od.Quantity,
        //                           Price = od.UnitPrice
        //                       });
        //    return View(orderDetail.ToList());
        //}
        public ActionResult Index(string dateFrom, string dateTo)
        {
            IQueryable<Orders> orderDetail;
            if (String.IsNullOrEmpty(dateFrom) || String.IsNullOrEmpty(dateTo))
            {
                 orderDetail = (from od in db.OrderDetail
                                   join o in db.Order on od.OrderID equals o.ID
                                   join p in db.Product on od.ProductID equals p.ID
                                   select new Orders
                                   {
                                       OrderID = od.OrderID,
                                       OrderDate = o.OrderDate,
                                       ProductId = od.ProductID,
                                       ProductName = p.Name,
                                       Quntity = od.Quantity,
                                       Price = od.UnitPrice
                                   }); 
            }
            else 
            {
                var dateF = DateTime.Parse(dateFrom);
                var dateT = DateTime.Parse(dateTo);
                orderDetail = (from od in db.OrderDetail
                               join o in db.Order on od.OrderID equals o.ID
                               join p in db.Product on od.ProductID equals p.ID
                               where o.OrderDate >= dateF && o.OrderDate <= dateT
                               select new Orders
                               {
                                   OrderID = od.OrderID,
                                   OrderDate = o.OrderDate,
                                   ProductId = od.ProductID,
                                   ProductName = p.Name,
                                   Quntity = od.Quantity,
                                   Price = od.UnitPrice
                               });
            }
            CreateExcell(orderDetail.ToList());
            return View(orderDetail.ToList());
        }
        [HttpPost]
        public void SendEmail(string emailAddress)
        {
            MailAddress from = new MailAddress("listorderstest@gmail.com", "Test");
            MailAddress to = new MailAddress(emailAddress);
            MailMessage message = new MailMessage(from, to);
            message.Subject = "Orders list";
            message.Body = "Auto-generated message<";
            message.IsBodyHtml = true;
            var path = Path.Combine(Server.MapPath("~/App_Data"), "Orders_" + DateTime.Now.ToShortDateString() + ".xls");
            message.Attachments.Add(new Attachment(path));
            SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587);
            smtp.Credentials = new NetworkCredential("listorderstest@gmail.com", "qaz123edc");
            smtp.EnableSsl = true;
            smtp.Send(message);
            smtp.Dispose();
            message.Dispose();
            Response.Redirect("SendSucces");
        }
        public static double GetDouble(string value)
        {
            double result;

            // Try parsing in the current culture
            if (!double.TryParse(value, System.Globalization.NumberStyles.Any, CultureInfo.CurrentCulture, out result) &&
                // Then try in US english
                !double.TryParse(value, System.Globalization.NumberStyles.Any, CultureInfo.GetCultureInfo("en-US"), out result) &&
                // Then in neutral language
                !double.TryParse(value, System.Globalization.NumberStyles.Any, CultureInfo.InvariantCulture, out result))
            { return result; }
            return result;
        }
        public void CreateExcell(List<Orders> data)
        {
            HSSFWorkbook wb;
            HSSFSheet sheetOrders;
            wb = new HSSFWorkbook();
            sheetOrders = (HSSFSheet)wb.CreateSheet("Orders");
            HSSFFont font = (HSSFFont)wb.CreateFont();
            HSSFFont fontС = (HSSFFont)wb.CreateFont();
            HSSFCellStyle styleHeader = (HSSFCellStyle)wb.CreateCellStyle();
            HSSFCellStyle styleBorder = (HSSFCellStyle)wb.CreateCellStyle();
            styleHeader.FillForegroundColor = IndexedColors.LightYellow.Index;
            styleHeader.FillPattern = FillPattern.SolidForeground;
            styleHeader.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            styleBorder.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            styleBorder.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            styleBorder.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            styleBorder.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            styleHeader.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            styleHeader.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            styleHeader.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            styleHeader.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            var currentRow = sheetOrders.CreateRow(0);
            var currentCell = currentRow.CreateCell(0);
            currentCell.SetCellValue("Orders list from" +data.OrderBy(date => date.OrderDate).First() + " to " + data.OrderBy(date => date.OrderDate).Last());
            currentRow = sheetOrders.CreateRow(1);
            currentCell = currentRow.CreateCell(0);
            currentCell.SetCellValue("Order id");
            currentCell.CellStyle = styleHeader;
            currentCell = currentRow.CreateCell(1);
            currentCell.SetCellValue("Order date");
            currentCell.CellStyle = styleHeader;
            currentCell = currentRow.CreateCell(2);
            currentCell.SetCellValue("Product id");
            currentCell.CellStyle = styleHeader;
            currentCell = currentRow.CreateCell(3);
            currentCell.SetCellValue("Product name");
            currentCell.CellStyle = styleHeader;
            currentCell = currentRow.CreateCell(4);
            currentCell.SetCellValue("Quantity");
            currentCell.CellStyle = styleHeader;
            currentCell = currentRow.CreateCell(5);
            currentCell.SetCellValue("Price");
            currentCell.CellStyle = styleHeader;
            currentCell = currentRow.CreateCell(6);
            currentCell.SetCellValue("Total price");
            currentCell.CellStyle = styleHeader;
            font.IsBold = true;
            font.IsItalic = true;
            HSSFCellStyle styleRightBorder = (HSSFCellStyle)wb.CreateCellStyle();
            int i = 1;
            foreach(var item in data)
            {
                i++;
                currentRow = sheetOrders.CreateRow(i);
                currentCell = currentRow.CreateCell(0);
                currentCell.CellStyle = styleBorder;
                currentCell.SetCellValue(item.OrderID.ToString());
                currentCell = currentRow.CreateCell(1);
                currentCell.CellStyle = styleBorder;
                currentCell.SetCellValue(item.OrderDate.ToString());
                currentCell = currentRow.CreateCell(2);
                currentCell.CellStyle = styleBorder;
                currentCell.SetCellValue(item.ProductId.ToString());
                currentCell = currentRow.CreateCell(3);
                currentCell.CellStyle = styleBorder;
                currentCell.SetCellValue(item.ProductName.ToString());
                currentCell = currentRow.CreateCell(4);
                currentCell.CellStyle = styleBorder;
                currentCell.SetCellValue(GetDouble(item.Quntity.ToString()));
                currentCell = currentRow.CreateCell(5);
                currentCell.CellStyle = styleBorder;
                currentCell.SetCellValue(GetDouble(item.Price.ToString()));
                currentCell = currentRow.CreateCell(6);
                currentCell.SetCellType(CellType.Formula);
                currentCell.CellStyle = styleBorder;
                currentCell.CellFormula=(String.Format("E{0}*F{0}",i+1));
            }
            for (i = 0; i < 7; i++)
            {
                sheetOrders.AutoSizeColumn(i);
            }
            //HSSFFormulaEvaluator.EvaluateAllFormulaCells(wb);
            var path = Path.Combine(Server.MapPath("~/App_Data"),"Orders_"+DateTime.Now.ToShortDateString()+ ".xls") ;
            using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                wb.Write(fs);
            };
        }
        public ActionResult SendSucces()
        {
            return View();
        }
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
