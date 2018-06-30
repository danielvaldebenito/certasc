<%@ WebHandler Language="C#" Class="Config" %>

using System;
using System.Web;
using System.Linq;
using System.Web.Script.Serialization;

public class Config : IHttpHandler {

    public void ProcessRequest (HttpContext context) {
        var post = HttpContext.Current;
        var data = new object();
        var action = post.Request["1"];
        var serializer = new JavaScriptSerializer();
        string sidx, sord;
        int page, rows;
        //config = new ConfigVariables();
        switch (action)
        {
            case "setMoneda":
                var type = post.Request["type"];
                var value = post.Request["value"];
                data = SetMoneda(type, value);
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
    private static object SetMoneda(string type, string value)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var config = db.Settings.Find(type);
                if (config == null)
                {
                    db.Settings.Add(new Settings  { Clave = type, Valor = value, Fecha = DateTime.Now });
                }
                else
                {
                    config.Valor = value;
                    config.Fecha = DateTime.Now;
                }
                db.SaveChanges();
                return new { done = true, message = "Moneda establecida" };
            }
        }
        catch (Exception ex)
        {
            return new { done = false, ex = ex.ToString() };
        }
    }
    public bool IsReusable {
        get {
            return true;
        }
    }

}