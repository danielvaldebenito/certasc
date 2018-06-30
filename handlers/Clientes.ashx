<%@ WebHandler Language="C#" Class="Clientes" %>

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

public class Clientes : IHttpHandler {

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
        switch (action)
        {
            case "grid":
                sidx = post.Request["sidx"];
                sord = post.Request["sord"];
                page = int.Parse(post.Request["page"]);
                rows = int.Parse(post.Request["rows"]);
                var name = post.Request["name"];
                var rut = post.Request["rut"];
                data = GetClientes(sidx, sord, page, rows, name, rut);
                break;
            case "enabledOrDisabled":
                var id = int.Parse(post.Request["id"]);
                data = EnableOrDisable(id);
                break;
            case "editClient":
                data = EditClient(post);
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
    private static object EditClient(HttpContext post)
    {
        try
        {
            var rut = post.Request["rut"];
            var nombre = post.Request["nombre"];
            var direccion = post.Request["direccion"];
            var telefono = post.Request["telefono"];
            var email = post.Request["email"];
            var giro = post.Request["giro"];
            var id = int.Parse(post.Request["id"]);
            using (var db = new CertelEntities())
            {
                var existClient = db.Cliente
                                    .Find(id);
                if (existClient == null)
                {
                    return new
                    {
                        done = false,
                        message = "Error, Cliente no existe",
                    };
                }
                existClient.Nombre = nombre;
                existClient.Direccion = direccion;
                existClient.EmailContacto = email;
                existClient.Rut = rut;
                existClient.TelefonoContacto = telefono;
                existClient.Giro = giro;
                db.SaveChanges();
                return new
                {
                    done = true,
                    message = "Cliente modificado correctamente"
                };

            }
        }
        catch (Exception ex)
        {
            log.Error("Excepción Clientes/EditClient", ex);
            return new
            {
                done = false,
                message = ex.ToString()
            };
        }
    }
    private static object EnableOrDisable (int id)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var client = db.Cliente.Find(id);
                if (client == null) return new { done = false, message = "Error: Cliente no existe" };

                client.Habilitado = !client.Habilitado;
                db.SaveChanges();
                var hab = client.Habilitado ? "habilitado" : "deshabilitado";
                return new { done = true, message = "Cliente " + hab + " exitosamente" };

            }
        }
        catch (Exception ex)
        {
            return new { done = false, ex = ex.Message };

        }
    }
    private static object GetClientes (string sidx, string sord, int page, int rows, string name, string rut)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var list = db.Cliente
                                .Where(w => w.Nombre.Contains(name))
                                .Where(w => w.Rut.Contains(rut));

                int pageSize = rows;
                int totalRecords = list.Count();
                int totalPages = (int)Math.Ceiling((float)totalRecords / (float)pageSize);
                totalPages = totalPages == 0 ? 1 : totalPages;
                var result = list
                    //.OrderBy(sidx + " " + sord)
                    .OrderByDescending(o => o.Habilitado)
                    .ThenBy(o => o.Nombre)
                    .Skip((page - 1) * pageSize)
                    .Take(pageSize);

                var grid = new
                {
                    page = page,
                    records = totalRecords,
                    total = totalPages,
                    rows = result
                           .Select(x => new
                           {
                               Id = x.ID,
                               Rut = x.Rut,
                               Nombre = x.Nombre == null ? string.Empty : x.Nombre.ToUpper(),
                               Direccion = x.Direccion == null ? string.Empty : x.Direccion.ToUpper(),
                               Telefono = x.TelefonoContacto,
                               Email = x.EmailContacto == null ? string.Empty : x.EmailContacto.ToLower(),
                               Giro = x.Giro,
                               Enabled = x.Habilitado
                           })
                           .ToList()
                };
                return grid;

            }
        }
        catch (Exception ex)
        {
            return new { done = false, ex = ex.Message };

        }
    }
    public bool IsReusable {
        get {
            return true;
        }
    }

}