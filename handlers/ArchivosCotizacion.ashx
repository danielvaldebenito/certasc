<%@ WebHandler Language="C#" Class="ArchivosCotizacion" %>


using System;
using System.Web;
using System.Linq;
using System.Web.Script.Serialization;
using System.Linq.Dynamic;
using System.Web.SessionState;
using System.Collections.Generic;
using System.IO;
using System.Drawing;
using System.Threading;

public class ArchivosCotizacion : IHttpHandler
{

    private static readonly log4net.ILog log = log4net.LogManager.GetLogger
  (System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
    public void ProcessRequest(HttpContext context)
    {

        var post = HttpContext.Current;
        var data = new object();
        var action = post.Request["1"];
        var serializer = new JavaScriptSerializer();
        switch (action)
        {
            case "getFiles":
                data = getFiles(post);
                break;

            case "updateQuotation":
                data = updateQuotation(post);
                break;
            case "dataQuotation":
                data = dataQuotation(post);
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

    private static object getFiles(HttpContext post)
    {
        try
        {
            var quotationId = int.Parse(post.Request["quotationId"]);
            using (var db = new CertelEntities())
            {

                var data = db.ArchivoCotizacion.Where(w => w.idCotizacion == quotationId)
                    .Select(s => new FilesData
                    {
                        id = s.Archivos.Id,
                        url = s.Archivos.URL,
                        nombre = s.Archivos.Nombre,
                        extencion = s.Archivos.Extension,
                        uso = s.Archivos.Uso,
                        isQuotation = false
                    }).ToList();

                var quotationFile = db.Cotizacion.Where(w => w.Id == quotationId).Select(s => new FilesData
                {
                    id = 0,
                    url = s.Filename,
                    nombre = s.Filename,
                    extencion = "pdf",
                    uso = "ARCHIVO_COTIZACION",
                    isQuotation = true
                }).FirstOrDefault();

                if (quotationFile.url != null)
                    data.Add(quotationFile);


                return new
                {
                    done = true,
                    data = data.OrderBy(o => o.id).ToList()
                };

            }
        }
        catch (Exception ex)
        {
            return new { done = false, ex = ex.Message };
        }
    }

    private static object updateQuotation(HttpContext post)
    {
        try
        {
            var quotationId = int.Parse(post.Request["quotationId"]);
            using (var db = new CertelEntities())
            {

                var data = db.Cotizacion.Where(w => w.Id == quotationId).FirstOrDefault();

                if (data != null)
                {
                    if (data.EstadoId == 1)
                    {
                        data.EstadoId = 2;

                        var hito = new Hito();
                        hito.Fecha = DateTime.Now;
                        hito.TipoHitoId = 16;
                        hito.Observacion = "La Cotizacion ha comenzado la evaluacion del cliente.";
                        hito.CotizacionId = quotationId;
                        db.Hito.Add(hito);


                        db.SaveChanges();

                        return new
                        {
                            done = true,
                            message = "Cotizacion Actualizada."
                        };
                    }
                    else
                    {
                        return new
                        {
                            done = true,
                            message = "La Cotizacion, ha sido actualizada anteriormente."
                        };
                    }

                }
                else
                {
                    return new
                    {
                        done = false,
                        message = "La cotizacion consultada, no existe."
                    };
                }


            }
        }
        catch (Exception ex)
        {
            return new { done = false, ex = ex.Message };
        }
    }



    public class FilesData
    {
        public int id { get; set; }
        public string url { get; set; }
        public string nombre { get; set; }
        public string extencion { get; set; }
        public string uso { get; set; }
        public bool isQuotation { get; set; }
    }

    private static object dataQuotation(HttpContext post)
    {
        try
        {
            var quotationId = int.Parse(post.Request["quotationId"]);
            using (var db = new CertelEntities())
            {

                var data = db.Cotizacion.Where(w => w.Id == quotationId)
                        .Select(s => new
                        {
                            IT = s.IT
                        })
                        .FirstOrDefault();


                return new
                {
                    done = true,
                    data = data
                };

            }
        }
        catch (Exception ex)
        {
            return new { done = false, ex = ex.Message };
        }
    }

    public bool IsReusable
    {
        get
        {
            return false;
        }
    }

}