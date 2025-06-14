using SistemaAuditoria.Models;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Web.Mvc;

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

            var sb = new System.Text.StringBuilder();
            sb.AppendLine("Proceso\tMarco Normativo\tNivel CMMI\tComentario\tFecha");

            foreach (var a in lista)
            {
                sb.AppendLine($"{a.NombreProceso}\t{a.MarcoNormativo}\t{a.NivelCMMI}\t{a.Comentario}\t{a.FechaRegistro.ToShortDateString()}");
            }

            Response.Clear();
            Response.AddHeader("content-disposition", "attachment;filename=auditorias.xls");
            Response.ContentType = "application/vnd.ms-excel";
            Response.ContentEncoding = System.Text.Encoding.UTF8;
            Response.Write(sb.ToString());
            Response.End();

            return null;
        }

    }
}
