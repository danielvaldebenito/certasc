<%@ WebHandler Language="C#" Class="Productos" %>

using System;
using System.Web;
using System.Linq;
using System.Web.Script.Serialization;
using System.Linq.Dynamic;

public class Productos : IHttpHandler {
    private static readonly log4net.ILog log = log4net.LogManager.GetLogger
  (System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
    public void ProcessRequest (HttpContext context) {
        var post = HttpContext.Current;
        var data = new object();
        var action = post.Request["1"];
        var serializer = new JavaScriptSerializer();

        //config = new ConfigVariables();
        switch (action)
        {
            case "grid":
                data = GetProducts (post);
                break;
            case "add":
                data = Save(post);
                break;
            case "edit":
                data = Edit(post);
                break;
            case "enableOrDisable":
                data = EnableOrDisable(post);
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

    private static object GetProducts (HttpContext post)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var sidx = post.Request["sidx"];
                var sord = post.Request["sord"];
                var page = int.Parse(post.Request["page"]);
                var rows = int.Parse(post.Request["rows"]);
                var name = post.Request["name"];

                var list = db.Producto
                                .Where(w => w.Descripcion.Contains(name));
                int pageSize = rows;
                int totalRecords = list.Count();
                int totalPages = (int)Math.Ceiling((float)totalRecords / (float)pageSize);
                totalPages = totalPages == 0 ? 1 : totalPages;
                var result = list
                    .OrderBy(sidx + " " + sord)
                    .Skip((page - 1) * pageSize)
                    .Take(pageSize);

                var grid = new
                {
                    page = page,
                    records = totalRecords,
                    total = totalPages,
                    rows = result
                           .Select(s => new
                           {
                               Id = s.Id,
                               Nombre = s.Descripcion,
                               Habilitado = s.Habilitado
                           })
                           .ToList()
                };
                return grid;
            }
        }
        catch (Exception ex)
        {
            return new { done = false, message = ex.ToString() };
        }
    }

    private static object Save (HttpContext post)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var name = post.Request["name"].Trim();

                var exists = db.Producto
                                .Where(w => w.Descripcion == name)
                                .Any();
                if(exists)
                {
                    return new { done = false, message = "Ya existe un producto llamado " + name };
                }

                var producto = new Producto
                {
                    Descripcion = name,
                    Habilitado = true
                };
                db.Producto.Add(producto);
                db.SaveChanges();
                producto.Codigo = producto.Id.ToString();
                db.SaveChanges();
                return new { done = true, message = "Registro ingresado correctamente" };
            }
        }
        catch (Exception ex)
        {
            return new { done = false, message = ex.ToString() };
        }
    }

    private static object EnableOrDisable (HttpContext post)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var id = int.Parse(post.Request["id"]);

                var producto = db.Producto
                                        .Find(id);
                var isEnabled = producto.Habilitado == true;
                producto.Habilitado = producto.Habilitado.HasValue ? !producto.Habilitado.Value : true;
                db.SaveChanges();
                return new { done = true, message = isEnabled ? "Registro deshabilitado correctamente" : "Registro habilitado correctamente" };
            }
        }
        catch (Exception ex)
        {
            return new { done = false, message = ex.ToString() };
        }
    }

    private static object Edit (HttpContext post)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var id = int.Parse(post.Request["id"]);
                var name = post.Request["name"];
                var exists = db.Producto
                                .Where(w => w.Descripcion == name)
                                .Where(w => w.Id != id)
                                .Any();
                if (exists)
                    return new { done = false, message = "Ya existe en el sistema, otro producto llamado " + name + ". Si está deshabilitado, puede volver a habilitarlo"};

                var product = db.Producto.Find(id);
                if (product == null)
                    return new { done = false, message = "Error, no existe producto" };

                product.Descripcion = name;
                db.SaveChanges();
                return new { done = true, message = "Registro editado correctamente" };
            }
        }
        catch (Exception ex)
        {
            return new { done = false, message = ex.ToString() };
        }
    }
    public bool IsReusable {
        get {
            return true;
        }
    }

}