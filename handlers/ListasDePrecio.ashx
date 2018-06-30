<%@ WebHandler Language="C#" Class="ListasDePrecio" %>

using System;
using System.Web;
using System.Linq;
using System.Web.Script.Serialization;
using System.Linq.Dynamic;

public class ListasDePrecio : IHttpHandler
{

    private static readonly log4net.ILog log = log4net.LogManager.GetLogger
          (System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
    public void ProcessRequest(HttpContext context)
    {
        var post = HttpContext.Current;
        var data = new object();
        var action = post.Request["1"];
        var serializer = new JavaScriptSerializer();
        string sidx, sord;
        int page, rows;
        switch (action)
        {
            case "grid":
                
                data = GetPriceList(post);
                break;

            case "addPriceList":
                data = addPriceList(post);
                break;

            case "modPriceList":
                data = modPriceList(post);
                break;

            case "delPriceList":
                data = delPriceList(post);
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

    private static object GetPriceList(HttpContext post)
    {
        try
        {
            var sidx = post.Request["sidx"];
            var sord = post.Request["sord"];
            var page = int.Parse(post.Request["page"]);
            var rows = int.Parse(post.Request["rows"]);
            var elevador = int.Parse(post.Request["elevador"]);
            var paradas = int.Parse(post.Request["paradas"]);
            var moreThanOne = bool.Parse(post.Request["moreThanOne"]);
            var producto = int.Parse(post.Request["product"]);
            using (var db = new CertelEntities())
            {
                var list = db.ListaPrecio
                            .Where(w => elevador == 0 ? true : w.TipoElevadorId == elevador)
                            .Where(w => paradas == 0 ? true : w.Paradas == paradas)
                            .Where(w => moreThanOne ? w.MasDe1Equipo == true : true)
                            .Where(w => producto == 0 ? true : w.ProductoId == producto);

                int pageSize = rows;
                int totalRecords = list.Count();
                int totalPages = (int)Math.Ceiling((float)totalRecords / (float)pageSize);
                totalPages = totalPages == 0 ? 1 : totalPages;
                var result = list
                    .OrderBy(sidx + " " + sord)
                    //.ThenBy(o => o.Fecha)
                    .Skip((page - 1) * pageSize)
                    .Take(pageSize);

                var grid = new
                {
                    page = page,
                    records = totalRecords,
                    total = totalPages,
                    rows = result
                    .AsEnumerable()
                           .Select(x => new
                           {
                               Id = x.Id,
                               Aparato = x.Aparato.Nombre,
                               AparatoID = x.TipoElevadorId,
                               Paradas = x.Paradas,
                               MasDeUnEquipo = (bool) x.MasDe1Equipo ? "Si" : "No",
                               TipoProducto = x.Producto.Descripcion,
                               TipoProductoID = x.ProductoId,
                               valor = x.ValorUF
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


    private static object addPriceList(HttpContext post)
    {
        try
        {
            var TipoElevadorId = int.Parse(post.Request["TipoElevadorId"]);
            var Paradas = int.Parse(post.Request["Paradas"]);
            var MasDe1Equipo = post.Request["MasDe1Equipo"];
            var ProductoId = int.Parse(post.Request["ProductoId"]);
            var ValorUF = post.Request["ValorUF"];

            using (var db = new CertelEntities())
            {
                var boolMasDe1Equipo = MasDe1Equipo == "on" ? true : false;

                var ExisteListaPrecio = db.ListaPrecio
                    .Where(w => w.TipoElevadorId == TipoElevadorId)
                    .Where(w => w.Paradas == Paradas)
                    .Where(w => w.MasDe1Equipo == boolMasDe1Equipo)
                    .Where(w => w.ProductoId == ProductoId)
                    .Any();


                if (ExisteListaPrecio)
                {
                    return new
                    {
                        done = false,
                        message = "La lista que esta intentando modificar ya existe.",
                    };
                }


                var NuevaLista = new ListaPrecio
                {
                    TipoElevadorId = TipoElevadorId,
                    Paradas = Paradas,
                    MasDe1Equipo = boolMasDe1Equipo,
                    ProductoId = ProductoId,
                    ValorUF = double.Parse(ValorUF, System.Globalization.CultureInfo.InvariantCulture)
                };

                db.ListaPrecio.Add(NuevaLista);
                db.SaveChanges();


                return new
                {
                    done = true,
                    message = "Nueva lista creada exitosamente.",
                };
            }
        }
        catch (Exception ex)
        {
            return new
            {
                done = false,
                message = ex.ToString(),
                code = 0
            };
        }
    }

    private static object modPriceList(HttpContext post)
    {
        try
        {
            var id = int.Parse(post.Request["id"]);
            var TipoElevadorId = int.Parse(post.Request["TipoElevadorId"]);
            var Paradas = int.Parse(post.Request["Paradas"]);
            var MasDe1Equipo = post.Request["MasDe1Equipo"];
            var ProductoId = int.Parse(post.Request["ProductoId"]);
            var ValorUF = post.Request["ValorUF"];

            using (var db = new CertelEntities())
            {

                var ListaPrecio = db.ListaPrecio.Where(w => w.Id == id).FirstOrDefault();
                ListaPrecio.ValorUF = double.Parse(ValorUF, System.Globalization.CultureInfo.InvariantCulture);
                db.SaveChanges();


                return new
                {
                    done = true,
                    message = "Nueva lista creada exitosamente." + ValorUF.ToString(),
                };
            }
        }
        catch (Exception ex)
        {
            return new
            {
                done = false,
                message = ex.ToString(),
                code = 0
            };
        }
    }

    private static object delPriceList(HttpContext post)
    {
        try
        {
            var id = int.Parse(post.Request["id"]);

            using (var db = new CertelEntities())
            {

                var ListaPrecio = db.ListaPrecio.Where(w => w.Id == id).FirstOrDefault();
                db.ListaPrecio.Remove(ListaPrecio);
                db.SaveChanges();


                return new
                {
                    done = true,
                    message = "Lista de precio eliminada exitosamente"
                };
            }
        }
        catch (Exception ex)
        {
            return new
            {
                done = false,
                message = ex.ToString(),
                code = 0
            };
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