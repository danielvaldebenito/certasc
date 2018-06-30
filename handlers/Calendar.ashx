<%@ WebHandler Language="C#" Class="Calendar" %>

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

public class Calendar : IHttpHandler
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
            case "getCalendarData":
                var ingeniero = post.Request["ingeniero"];
                data = GetCalendarData(ingeniero);
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

    private static object GetCalendarData(string ingeniero)
    {
        try
        {
            using (var db = new CertelEntities())
            {

                var calendarData = db.Inspeccion
                           .Where(w => ingeniero == "0" ? true : w.Ingeniero == ingeniero)
                           .Where(w => w.EstadoID == 1)
                           .GroupBy(g => new
                           {
                               g.Servicio,
                               g.FechaInspeccion,
                               g.Ubicacion,
                               g.Usuario
                           })
                           .AsEnumerable()
                           .Select(s => new
                           {
                               title = s.Key.Ubicacion + " (" +  s.Count().ToString() + ")",
                               start = s.Key.FechaInspeccion == null ? "" : ((DateTime)s.Key.FechaInspeccion).ToString(("yyyy-MM-dd HH:mm:ss")),
                               ingeniero = s.Key.Usuario.Nombre + " " + s.Key.Usuario.Apellido
                           }).ToList();



                return new
                {
                    done = true,
                    data = calendarData
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
            return true;
        }
    }

}