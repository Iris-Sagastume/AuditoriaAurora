using SistemaAuditoria.Models;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Web.Mvc;
using OfficeOpenXml;


namespace SistemaAuditoria.Controllers
{
    public class AuditoriaController : Controller
    {
        string connectionString = "Data Source=ANTE-PC;Initial Catalog=BD;Integrated Security=True";


        public ActionResult Index()
        {
            List<Auditoria> lista = new List<Auditoria>();
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                string query = "SELECT * FROM Auditorias";
                SqlCommand cmd = new SqlCommand(query, con);
                con.Open();
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    lista.Add(new Auditoria
                    {
                        Id = Convert.ToInt32(rdr["Id"]),
                        NombreProceso = rdr["NombreProceso"].ToString(),
                        MarcoNormativo = rdr["MarcoNormativo"].ToString(),
                        NivelCMMI = Convert.ToInt32(rdr["NivelCMMI"]),
                        Comentario = rdr["Comentario"].ToString(),
                        FechaRegistro = Convert.ToDateTime(rdr["FechaRegistro"])
                    });
                }
            }
            return View(lista);
        }

        public ActionResult Crear()
        {
            return View();
        }
        public ActionResult HtmlPage1()
        {
            return View();
        }
        public ActionResult HtmlPage2()
        {
            return View();
        }
        public ActionResult HtmlPage3()
        {
            return View();
        }
        public ActionResult HtmlPage4()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Crear(Auditoria a)
        {
            if (ModelState.IsValid)
            {
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    string query = "INSERT INTO Auditorias (NombreProceso, MarcoNormativo, NivelCMMI, Comentario) VALUES (@NombreProceso, @MarcoNormativo, @NivelCMMI, @Comentario)";
                    SqlCommand cmd = new SqlCommand(query, con);
                    cmd.Parameters.AddWithValue("@NombreProceso", a.NombreProceso);
                    cmd.Parameters.AddWithValue("@MarcoNormativo", a.MarcoNormativo);
                    cmd.Parameters.AddWithValue("@NivelCMMI", a.NivelCMMI);
                    cmd.Parameters.AddWithValue("@Comentario", a.Comentario);
                    con.Open();
                    cmd.ExecuteNonQuery();
                }
                return RedirectToAction("Index");
            }
            return View(a);
        }

        public ActionResult ExportarExcel()
        {
                       var lista = new List<Auditoria>();
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                string query = "SELECT * FROM Auditorias";
                SqlCommand cmd = new SqlCommand(query, con);
                con.Open();
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    lista.Add(new Auditoria
                    {
                        NombreProceso = rdr["NombreProceso"].ToString(),
                        MarcoNormativo = rdr["MarcoNormativo"].ToString(),
                        NivelCMMI = Convert.ToInt32(rdr["NivelCMMI"]),
                        Comentario = rdr["Comentario"].ToString(),
                        FechaRegistro = Convert.ToDateTime(rdr["FechaRegistro"])
                    });
                }
            }


            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Auditorias");

                // Encabezados
                ws.Cells[1, 1].Value = "Proceso";
                ws.Cells[1, 2].Value = "Marco Normativo";
                ws.Cells[1, 3].Value = "Nivel CMMI";
                ws.Cells[1, 4].Value = "Comentario";
                ws.Cells[1, 5].Value = "Fecha";

                int row = 2;
                foreach (var item in lista)
                {
                    ws.Cells[row, 1].Value = item.NombreProceso;
                    ws.Cells[row, 2].Value = item.MarcoNormativo;
                    ws.Cells[row, 3].Value = item.NivelCMMI;
                    ws.Cells[row, 4].Value = item.Comentario;
                    ws.Cells[row, 5].Value = item.FechaRegistro.ToShortDateString();
                    row++;
                }

                ws.Cells[1, 1, row - 1, 5].AutoFitColumns();
                ws.Cells["A1:E1"].Style.Font.Bold = true;

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=auditorias.xlsx");
                Response.BinaryWrite(package.GetAsByteArray());
                Response.End();
            }

            return null;
        }

    }
}
