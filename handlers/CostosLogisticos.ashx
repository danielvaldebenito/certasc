<%@ WebHandler Language="C#" Class="Cotizaciones" %>

using System;
using System.Web;
using System.Linq;
using System.Web.Script.Serialization;
using System.Linq.Dynamic;

public class Cotizaciones : IHttpHandler {
    private static readonly log4net.ILog log = log4net.LogManager.GetLogger
   (System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
    //public static ConfigVariables config { get; set; }
    public void ProcessRequest (HttpContext context) {
        log4net.Config.DOMConfigurator.Configure();
        var post = HttpContext.Current;
        var data = new object();
        var action = post.Request["1"];
        var serializer = new JavaScriptSerializer();

        //config = new ConfigVariables();
        switch (action)
        {
            case "grid-costos-logisticos":
                data = GetCostosLogisticos(post);
                break;
            case "add":
                data = Save(post);
                break;
            case "edit":
                data = Edit(post);
                break;
            case "delete":
                data = Delete(post);
                break;
            case "addType":
                data = AddType(post);
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
    private static object AddType(HttpContext post)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var name = post.Request["name"];
                var exists = db.TipoGasto
                                .Where(w => w.Descripcion == name)
                                .Any();
                if (exists)
                    return new { done = false, message = "Ya existe un tipo de costo operacional llamado " + name };
                var newId = db.TipoGasto.Max(m => m.Id) + 1;
                var tg = new TipoGasto
                {
                    Descripcion =  name,
                    Id = newId
                };
                db.TipoGasto.Add(tg);
                db.SaveChanges();
                return new { done = true, message = "Tipo de  gasto agregado exitosamente", id = tg.Id };
            }
        }
        catch (Exception ex)
        {
            return new { done = false, ex = ex.ToString(), message = "Ha ocurrido un error" };
        }
    }
    private static object GetCostosLogisticos (HttpContext post)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var sidx = post.Request["sidx"];
                var sord = post.Request["sord"];
                var page = int.Parse(post.Request["page"]);
                var rows = int.Parse(post.Request["rows"]);
                var ciudad = int.Parse(post.Request["ciudad"]);
                var tipoCosto = int.Parse(post.Request["tipoCosto"]);
                var list = db.ListaPrecioCostosLogisticos
                                .Where(w => ciudad == 0 ? true : w.CiudadId == ciudad)
                                .Where(w => tipoCosto == 0 ? true : w.TipoGastoId == tipoCosto);
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
                               Ciudad = s.Ciudad.Descripcion,
                               CiudadId = s.CiudadId,
                               TipoCostoId = s.TipoGastoId,
                               TipoGasto = s.TipoGasto.Descripcion,
                               Valor = s.ValorCLP
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
                var ciudad = int.Parse(post.Request["ciudad"]);
                var tipoGasto = int.Parse(post.Request["tipoGasto"]);
                var valor = double.Parse(post.Request["valor"], System.Globalization.CultureInfo.CurrentCulture);

                var exists = db.ListaPrecioCostosLogisticos
                                .Where(w => w.CiudadId == ciudad)
                                .Where(w => w.TipoGastoId == tipoGasto)
                                .Any();
                if(exists)
                {
                    return new { done = false, message = "Ya existe un valor para la ciudad y tipo de gasto seleccionado" };
                }

                var costoLogistico = new ListaPrecioCostosLogisticos
                {
                    CiudadId = ciudad,
                    TipoGastoId = tipoGasto,
                    ValorCLP = valor
                };
                db.ListaPrecioCostosLogisticos.Add(costoLogistico);
                db.SaveChanges();
                return new { done = true, message = "Registro ingresado correctamente" };
            }
        }
        catch (Exception ex)
        {
            return new { done = false, message = ex.ToString() };
        }
    }

    private static object Delete (HttpContext post)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var id = int.Parse(post.Request["id"]);

                var costoLogistico = db.ListaPrecioCostosLogisticos
                                        .Find(id);
                db.ListaPrecioCostosLogisticos.Remove(costoLogistico);
                db.SaveChanges();
                return new { done = true, message = "Registro eliminado correctamente" };
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
                var valor = double.Parse(post.Request["valor"], System.Globalization.CultureInfo.CurrentCulture);
                var costoLogistico = db.ListaPrecioCostosLogisticos
                                        .Find(id);
                costoLogistico.ValorCLP = valor;
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