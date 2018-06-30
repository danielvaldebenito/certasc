<%@ WebHandler Language="C#" Class="SetInforme" %>

using System;
using System.Web;
using System.Linq;
using System.Web.Script.Serialization;
using System.Linq.Dynamic;
using System.Web.SessionState;
using System.Collections.Generic;
using System.IO;
using System.Globalization;
public class SetInforme : IHttpHandler, IRequiresSessionState {

    private static readonly log4net.ILog log = log4net.LogManager.GetLogger
  (System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
    public void ProcessRequest (HttpContext context) {
        log4net.Config.DOMConfigurator.Configure();
        var post = HttpContext.Current;
        var data = new object();
        var action = post.Request["1"];
        var serializer = new JavaScriptSerializer();
        string sidx, sord;
        int page, rows;
        switch(action)
        {
            case "getPlantillaAlcance":
                data = GetPlantillaAlcance(post);
                break;
            case "createPdf":
                data = CreatePdf(post);
                break;
            case "start":
                data = StartPdf(post);
                break;
            case "revisar":
                data = Revisar(post);
                break;
            case "aprobar":
                data = Aprobar(post);
                break;
        }
        if (data != null)
        {
            var json = serializer.Serialize(data);
            context.Response.ContentType = "application/json";
            context.Response.Write(json);
            context.Response.Flush();

        }
    }
    private static object Plantilla(HttpContext post)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                return null;
            }
        }
        catch (Exception ex)
        {
            log.Error("SetInforme/Plantilla", ex);
            return new { done = false, message = ex.ToString() };
        }
    }

    private static object Aprobar(HttpContext post)
    {
        try
        {
            var id = int.Parse(post.Request["id"]);
            var fecha = DateTime.ParseExact(post.Request["fecha"], "dd-MM-yyyy", CultureInfo.InvariantCulture);
            var destinatario = post.Request["destinatario"];
            using (var db = new CertelEntities())
            {
                var inspeccion = db.Inspeccion
                                    .Find(id);
                if (inspeccion == null) return new { done = false, message = "Error: Inspección no existe" };

                inspeccion.Destinatario = destinatario;
                inspeccion.FechaEntrega = fecha;

                if(inspeccion.FechaRevision == null)
                {
                    Revisar(post);
                }
                var fase2 = inspeccion.Inspeccion1;
                foreach(var f in fase2)
                {
                    f.FechaInspeccion = fecha.AddDays((int)inspeccion.DiasPlazo);
                }
                db.SaveChanges();
                return new { done = true, message = "Datos ingresados correctamente!" };
            }
        }
        catch (Exception ex)
        {
            log.Error("SetInforme/Aprobar", ex);
            return new { done = false, message = ex.ToString() };
        }
    }
    private static object Revisar(HttpContext post)
    {
        try
        {
            var id = int.Parse(post.Request["id"]);
            var user = post.Request["user"];
            var usuario = new Encriptacion(user, false).newText;
            using (var db = new CertelEntities())
            {
                var inspeccion = db.Inspeccion
                                    .Find(id);
                if (inspeccion == null) return new { done = false, message = "Error" };
                if(inspeccion.InspeccionNorma.Count == 0) return new { done = false, message = "No hay normas asociadas" };
                inspeccion.FechaRevision = DateTime.Now;
                inspeccion.Revisador = usuario;
                db.SaveChanges();
                return new { done = true, message = "Inspección ha sido revisada" };
            }
        }
        catch (Exception ex)
        {
            log.Error("SetInforme/Revisar", ex);
            return new { done = false, message = ex.ToString() };
        }
    }
    private static object StartPdf(HttpContext post)
    {
        try
        {
            var id = int.Parse(post.Request["id"]);
            using (var db = new CertelEntities())
            {
                var inspeccion = db.Inspeccion
                                    .Where(w => w.ID == id)
                                    .FirstOrDefault();
                if (inspeccion == null)
                    return new { done = false, message = "Inspeccion no existe" };

                var informe = inspeccion.Informe.FirstOrDefault();
                if(informe == null)
                {
                    return new { done = false, message = "El informe no está accesible" };
                }
                var pdf = informe.FileName;
                return new { done = true, message = "Presione OK para ver el informe", url = pdf };
            }
        }
        catch (Exception ex)
        {
            log.Error("SetInforme/Plantilla", ex);
            return new { done = false, message = ex.ToString() };
        }
    }
    private static object CreatePdf(HttpContext post)
    {
        try
        {
            var id = int.Parse(post.Request["id"]);
            using (var db = new CertelEntities())
            {
                var inspeccion = db.Inspeccion
                                    .Find(id);
                if(inspeccion == null)
                    return new { done = false, message = "Inspeccion no existe" };

                if(inspeccion.TipoInformeID == null)
                    return new { done = false, message = "Debe seleccionar un tipo de informe para la inspección" };
                if(!inspeccion.Cumplimiento.Any())
                    return new { done = false, message = "Asegúrese de llenar previamente el check-list correspondiente" };
                db.Database.CommandTimeout = 300;
                var path = string.Empty;
                switch(inspeccion.TipoInformeID)
                {
                    case 1:
                        var pdf = new CreatePDFD114(inspeccion);
                        path = pdf.Rendered;
                        break;
                    case 2:
                        var pdf2 = new CreatePDF4401(inspeccion);
                        path = pdf2.Rendered;
                        break;
                    case 3:
                        var pdf3 = new CreatePDFD116(inspeccion);
                        path = pdf3.Rendered;
                        break;
                    case 4:
                        var pdf4 = new CreatePDFD115(inspeccion);
                        path = pdf4.Rendered;
                        break;
                    case 5:
                        var pdf5 = new CreatePDFD112(inspeccion);
                        path = pdf5.Rendered;
                        break;
                    case 6:
                        var pdf6 = new CreatePDFD118(inspeccion);
                        path = pdf6.Rendered;
                        break;
                    case null:
                        path = "";
                        return new { done = false, message = "No ha especificado el tipo de informe que desea generar" };

                    default: return new { done = false, message = "Tipo de informe aún no ha sido creado" };

                }
                RemovePdfs();
                return new { done = true, message = "Informe generado exitosamente", path = path };;

            }
        }
        catch (Exception ex)
        {
            log.Error("SetInforme/Plantilla", ex);
            return new { done = false, message = ex.ToString() };
        }
    }
    private static bool RemovePdfs()
    {
        var pdfLifeHours = 48;
        string basePath = HttpContext.Current.Server.MapPath("~/pdf/");
        var dir = new DirectoryInfo(basePath);
        FileInfo[] files = dir.GetFiles();
        for(int i = 0; i < files.Length; i++)
        {
            FileInfo file = files[i];
            if(file.Exists)
            {
                if((DateTime.Now - file.CreationTime).TotalHours > pdfLifeHours)
                {
                    if(File.Exists(basePath + file.Name))
                        File.Delete(basePath + file.Name);
                }
            }
        }
        return true;
    }
    private static object GetPlantillaAlcance(HttpContext post)
    {
        try
        {
            var id = int.Parse(post.Request["id"]);
            using (var db = new CertelEntities())
            {
                var inspeccion = db.Inspeccion
                                    .Where(w => w.ID == id)
                                    .FirstOrDefault();
                var i = inspeccion.Aparato.Nombre;
                var normas = inspeccion.InspeccionNorma.Select(s => s.Norma.Nombre).ToList();
                string n = string.Empty;
                var count = 0;
                foreach(var nor in normas)
                {
                    count++;
                    var separador = ", ";
                    if (count == normas.Count - 1)
                        separador = " y ";
                    else if (count == normas.Count)
                        separador = string.Empty;
                    else
                        separador = ", ";
                    n += nor + separador;
                }
                var alcance = db.EstructuraInfome
                                .Where(w => w.ID == 1)
                                .AsEnumerable()
                                .Select(s => new
                                {
                                    Id = s.ID,
                                    Titulo = s.Título,
                                    Text = string.Format(s.Texto, i, n, i)
                                }).ToList();

                return alcance;
            }
        }
        catch (Exception ex)
        {
            log.Error("SetInforme/GetPlantillaAlcance", ex);
            return new { done = false, message = ex.ToString() };
        }
    }

    //private static object GetPlantillaIntro(HttpContext post)
    //{
    //    try
    //    {
    //        var id = int.Parse(post.Request["id"]);
    //        using (var db = new CertelEntities())
    //        {
    //            var inspeccion = db.Inspeccion
    //                                .Where(w => w.ID == id)
    //                                .FirstOrDefault();
    //            var i = inspeccion.Aparato.Nombre;
    //            var normas = inspeccion.InspeccionNorma.Select(s => s.Norma.Nombre).ToList();
    //            string n = string.Empty;
    //            var count = 0;
    //            foreach (var nor in normas)
    //            {
    //                count++;
    //                var separador = ", ";
    //                if (count == normas.Count - 1)
    //                    separador = " y ";
    //                else if (count == normas.Count)
    //                    separador = string.Empty;
    //                else
    //                    separador = ", ";
    //                n += nor + separador;
    //            }
    //            var intro = db.EstructuraInfome
    //                            .Where(w => w.ID == 1)
    //                            .AsEnumerable()
    //                            .Select(s => new
    //                            {
    //                                Id = s.ID,
    //                                Titulo = s.Título,
    //                                Text = string.Format(s.Texto, i, n, i)
    //                            }).ToList();

    //            return alcance;
    //        }
    //    }
    //    catch (Exception ex)
    //    {
    //        log.Error("SetInforme/GetPlantillaAlcance", ex);
    //        return new { done = false, message = ex.ToString() };
    //    }
    //}
    public bool IsReusable {
        get {
            return true;
        }
    }

}