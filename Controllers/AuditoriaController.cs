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
            var auditorias = new List<Auditoria>();
            var recomendaciones = new List<Recomendacion>();

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();

                // Auditorías
                var cmd1 = new SqlCommand("SELECT * FROM Auditorias", con);
                var rdr1 = cmd1.ExecuteReader();
                while (rdr1.Read())
                {
                    auditorias.Add(new Auditoria
                    {
                        Id = Convert.ToInt32(rdr1["Id"]),
                        NombreProceso = rdr1["NombreProceso"].ToString(),
                        MarcoNormativo = rdr1["MarcoNormativo"].ToString(),
                        NivelCMMI = Convert.ToInt32(rdr1["NivelCMMI"]),
                        Comentario = rdr1["Comentario"].ToString(),
                        FechaRegistro = Convert.ToDateTime(rdr1["FechaRegistro"])
                    });
                }
                rdr1.Close();

                // Recomendaciones
                var cmd2 = new SqlCommand("SELECT * FROM Recomendacion", con);
                var rdr2 = cmd2.ExecuteReader();
                while (rdr2.Read())
                {
                    recomendaciones.Add(new Recomendacion
                    {
                        Id = Convert.ToInt32(rdr2["Id"]),
                        AspectoEvaluado = rdr2["AspectoEvaluado"].ToString(),
                        Observaciones = rdr2["Observaciones"].ToString(),
                        RecomendacionesTexto = rdr2["Recomendaciones"].ToString(),
                        Riesgos = rdr2["Riesgos"].ToString(),
                        Fortalezas = rdr2["Fortalezas"].ToString(),
                        FechaRegistro = Convert.ToDateTime(rdr2["FechaRegistro"])
                    });
                }
            }

            return View(Tuple.Create(auditorias, recomendaciones));
        }


        public ActionResult Crear()
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

        // ----------------------------- //
        //   Métodos de Recomendaciones  //
        // ----------------------------- //

        public ActionResult GuardarRecomendacion(Recomendacion r)
        {
            if (ModelState.IsValid)
            {
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    string query = @"INSERT INTO Recomendacion (AspectoEvaluado, Observaciones, Recomendaciones, Riesgos, Fortalezas)
                             VALUES (@AspectoEvaluado, @Observaciones, @Recomendaciones, @Riesgos, @Fortalezas)";
                    SqlCommand cmd = new SqlCommand(query, con);
                    cmd.Parameters.AddWithValue("@AspectoEvaluado", r.AspectoEvaluado);
                    cmd.Parameters.AddWithValue("@Observaciones", r.Observaciones);
                    cmd.Parameters.AddWithValue("@Recomendaciones", r.RecomendacionesTexto);
                    cmd.Parameters.AddWithValue("@Riesgos", r.Riesgos);
                    cmd.Parameters.AddWithValue("@Fortalezas", r.Fortalezas);

                    con.Open();
                    cmd.ExecuteNonQuery();
                }
                return RedirectToAction("Index");
            }

            return View("HtmlPage1", r);
        }
        public ActionResult ExportarRecomendacionesExcel()
        {
            var lista = new List<Recomendacion>();

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                string query = "SELECT * FROM Recomendacion";
                SqlCommand cmd = new SqlCommand(query, con);
                con.Open();
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    lista.Add(new Recomendacion
                    {
                        AspectoEvaluado = rdr["AspectoEvaluado"].ToString(),
                        Observaciones = rdr["Observaciones"].ToString(),
                        RecomendacionesTexto = rdr["Recomendaciones"].ToString(),
                        Riesgos = rdr["Riesgos"].ToString(),
                        Fortalezas = rdr["Fortalezas"].ToString(),
                        FechaRegistro = Convert.ToDateTime(rdr["FechaRegistro"])
                    });
                }
            }

            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Recomendaciones");

                ws.Cells[1, 1].Value = "Aspecto Evaluado";
                ws.Cells[1, 2].Value = "Observaciones";
                ws.Cells[1, 3].Value = "Recomendaciones";
                ws.Cells[1, 4].Value = "Riesgos";
                ws.Cells[1, 5].Value = "Fortalezas";
                ws.Cells[1, 6].Value = "Fecha";

                int row = 2;
                foreach (var item in lista)
                {
                    ws.Cells[row, 1].Value = item.AspectoEvaluado;
                    ws.Cells[row, 2].Value = item.Observaciones;
                    ws.Cells[row, 3].Value = item.RecomendacionesTexto;
                    ws.Cells[row, 4].Value = item.Riesgos;
                    ws.Cells[row, 5].Value = item.Fortalezas;
                    ws.Cells[row, 6].Value = item.FechaRegistro.ToShortDateString();
                    row++;
                }

                ws.Cells[1, 1, row - 1, 6].AutoFitColumns();
                ws.Cells["A1:F1"].Style.Font.Bold = true;

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=recomendaciones.xlsx");
                Response.BinaryWrite(package.GetAsByteArray());
                Response.End();
            }

            return null;
        }


    }
}
