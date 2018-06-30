<%@ WebHandler Language="C#" Class="Combobox" %>

using System;
using System.Web;
using System.Linq;
using System.Web.Script.Serialization;
using System.Collections.Generic;
public class Combobox : IHttpHandler
{

    private static readonly log4net.ILog log = log4net.LogManager.GetLogger
   (System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
    public void ProcessRequest(HttpContext context)
    {

        log4net.Config.DOMConfigurator.Configure();
        var post = HttpContext.Current;
        var data = new object();
        var action = post.Request["1"];
        var serializer = new JavaScriptSerializer();

        switch (action)
        {
            case "tipoNorma":
                data = GetTipoNorma();
                break;
            case "clientes":
                data = GetClientes();
                break;
            case "aparatos":
                data = GetAparatos();
                break;
            case "inspectores":
                data = GetInspectores();
                break;
            case "tipoFuncionamiento":
                data = GetTipoFuncionamiento();
                break;
            case "normas":
                var norma = int.Parse(post.Request["norma"]);
                data = GetNormas(norma);
                break;
            case "tipoInforme":
                data = GetTipoInforme();
                break;
            case "destinoProyecto":
                data = GetDestinoProyecto();
                break;
            case "normasAsociadas":
                data = GetNormasInspeccion(post);
                break;
            case "titulosInspeccion":
                data = GetTitulosInspeccion(post);
                break;
            case "sameIt":
                data = GetItsBrothers(post);
                break;
            case "roles":
                data = GetRoles(post);
                break;
            case "monedas":
                data = GetMonedas(post);
                break;
            case "formasDePago":
                data = GetFormasDePago(post);
                break;
            case "vendedores":
                data = GetVendedores(post);
                break;
            case "hitos":
                data = GetHitos(post);
                break;
            case "estadosCotizacion":
                data = GetEstadosCotizacion();
                break;
            case "marcas":
                data = GetMarcas();
                break;
            case "ingenieros":
                data = GetIngenieros();
                break;
            case "tipoGasto":
                data = GetTipoGastos();
                break;
            case "productos":
                data = GetProductos();
                break;
            case "ciudades":
                data = GetCiudades();
                break;
            default:
                context.Response.Write("No se ha seleccionado ninguna opción");
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

    private static object GetCiudades()
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var ciudades = db.Ciudad
                                .OrderBy(o => o.Descripcion)
                                .Select(s => new
                                {
                                    Value = s.Id,
                                    Text = s.Descripcion
                                })
                                
                                .ToList();
                return new
                {
                    done = true,
                    message = "OK",
                    data = ciudades
                };
            }
        }
        catch (Exception ex)
        {
            return new { done = false, message = ex.ToString() };
        }
    }
    private static object GetEstadosCotizacion()
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var estados = db.EstadoCotizacion
                                .Select(s => new
                                {
                                    Value = s.Id,
                                    Text = s.Descripcion
                                })
                                .ToList();
                return new
                {
                    done = true,
                    message = "OK",
                    data = estados
                };

            }
        }
        catch (Exception ex)
        {
            return new { done = false, message = ex.ToString() };
        }
    }
    private static object GetProductos()
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var productos = db.Producto
                                .Where(w => w.Habilitado == true)
                                .Select(s => new
                                {
                                    Value = s.Codigo,
                                    Text = s.Descripcion
                                })
                                .ToList();
                return new
                {
                    done = true,
                    message = "OK",
                    data = productos
                };

            }
        }
        catch (Exception ex)
        {
            return new { done = false, message = ex.ToString() };
        }
    }
    private static object GetTipoGastos()
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var tipoGastos = db.TipoGasto
                                .Select(s => new
                                {
                                    Value = s.Id,
                                    Text = s.Descripcion
                                })
                                .ToList();
                return new
                {
                    done = true,
                    message = "OK",
                    data = tipoGastos
                };

            }
        }
        catch (Exception ex)
        {
            return new { done = false, message = ex.ToString() };
        }
    }
    private static object GetMarcas()
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var marcas = db.Marca
                                .Select(s => new
                                {
                                    Value = s.Id,
                                    Text = s.Descripcion
                                })
                                .ToList();
                return new
                {
                    done = true,
                    message = "OK",
                    data = marcas
                };

            }
        }
        catch (Exception ex)
        {
            return new { done = false, message = ex.ToString() };
        }
    }

    private static object GetIngenieros()
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var marcas = db.UsuarioRol.
                                Where(w => w.Rol == 1)
                                .Select(s => new
                                {
                                    Value = s.Usuario1.NombreUsuario,
                                    Text = s.Usuario1.Nombre + " " + s.Usuario1.Apellido
                                })
                                .ToList();
                return new
                {
                    done = true,
                    message = "OK",
                    data = marcas
                };

            }
        }
        catch (Exception ex)
        {
            return new { done = false, message = ex.ToString() };
        }
    }

    private static object GetMonedas(HttpContext post)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var monedas = db.Moneda
                            .OrderByDescending(o => o.Descripcion)
                            .Select(s => new
                            {
                                Value = s.Id,
                                Text = s.Descripcion
                            })
                            .ToList();
                return new
                {
                    done = true,
                    message = "OK",
                    data = monedas
                };
            }
        }
        catch (Exception ex)
        {
            log.Error("Excepción Combobox/Roles", ex);
            return new { done = false, message = ex.ToString() };
        }
    }

    private static object GetFormasDePago(HttpContext post)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var formasDePago = db.FormaDePago
                            .Select(s => new
                            {
                                Value = s.Id,
                                Text = s.Descripcion
                            })
                            .ToList();
                return new
                {
                    done = true,
                    message = "OK",
                    data = formasDePago
                };
            }
        }
        catch (Exception ex)
        {
            log.Error("Excepción Combobox/Roles", ex);
            return new { done = false, message = ex.ToString() };
        }
    }
    private static object GetVendedores(HttpContext post)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var vendedores = db.Usuario
                                .Where(w => w.UsuarioRol.Any(a => a.Rol == 4)) // Vendedor
                                .Select(s => new
                                {
                                    Value = s.NombreUsuario,
                                    Text = string.Concat(s.Nombre, " ", s.Apellido)
                                })
                                .ToList();
                return new
                {
                    done = true,
                    message = "OK",
                    data = vendedores
                };
            }
        }
        catch (Exception ex)
        {
            log.Error("Excepción Combobox/Roles", ex);
            return new { done = false, message = ex.ToString() };
        }
    }
    private static object GetItsBrothers(HttpContext post)
    {
        try
        {
            var inspeccionId = int.Parse(post.Request["id"]);
            using (var db = new CertelEntities())
            {
                var inspeccion = db.Inspeccion.Find(inspeccionId);
                var brothers = db.Inspeccion
                                    .Where(w => w.ServicioID == inspeccion.ServicioID)
                                    .Where(w => w.ID != inspeccionId)
                                    .Where(w => w.Fase == inspeccion.Fase);

                var data = brothers
                            .Distinct()
                            .Select(s => new
                            {
                                Value = s.ID,
                                Text = s.IT
                            })
                            .ToList();
                return new
                {
                    done = true,
                    message = "OK",
                    data = data
                };
            }
        }
        catch (Exception ex)
        {
            log.Error("Excepción Combobox/TitulosNorma", ex);
            return new { done = false, message = ex.ToString() };
        }
    }
    private static object GetRoles(HttpContext post)
    {
        try
        {
            var username = post.Request["username"];
            using (var db = new CertelEntities())
            {
                var roles = db.Rol
                            .Where(w => !w.UsuarioRol.Any(a => a.Usuario == username))
                            .Select(s => new
                            {
                                Value = s.ID,
                                Text = s.Nombre
                            })
                            .ToList();
                return new
                {
                    done = true,
                    message = "OK",
                    data = roles
                };
            }
        }
        catch (Exception ex)
        {
            log.Error("Excepción Combobox/Roles", ex);
            return new { done = false, message = ex.ToString() };
        }
    }
    private static object GetTitulosInspeccion(HttpContext post)
    {
        try
        {
            var inspeccionId = int.Parse(post.Request["inspeccionId"]);
            using (var db = new CertelEntities())
            {
                var inspeccion = db.Inspeccion.Find(inspeccionId);
                var titles = new TitulosInspeccion(inspeccion).ListaInspeccion;

                var data = titles
                            .Distinct()
                            .Select(s => new
                            {
                                Value = s.ID,
                                Text = s.Texto
                            })
                            .ToList();
                return new
                {
                    done = true,
                    message = "OK",
                    data = data
                };
            }
        }
        catch (Exception ex)
        {
            log.Error("Excepción Combobox/TitulosNorma", ex);
            return new { done = false, message = ex.ToString() };
        }
    }

    private static object GetNormasInspeccion(HttpContext post)
    {
        try
        {
            var inspeccionId = int.Parse(post.Request["inspeccionId"]);
            using (var db = new CertelEntities())
            {
                var data = db.InspeccionNorma
                            .Where(w => w.InspeccionID == inspeccionId)
                            .Distinct()
                            .OrderBy(o => o.Norma.NormasAsociadas1.Any())
                            .Select(s => new
                            {
                                Value = s.NormaID,
                                Text = s.Norma.Nombre
                            })
                            .ToList();
                return new
                {
                    done = true,
                    message = "OK",
                    data = data
                };
            }
        }
        catch (Exception ex)
        {
            log.Error("Excepción Combobox/NormasInspeccion", ex);
            return new { done = false, message = ex.ToString() };
        }
    }

    private static object GetDestinoProyecto()
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var data = db.DestinoProyecto
                            .OrderBy(o => o.Descripcion)
                            .Select(s => new
                            {
                                Value = s.ID,
                                Text = s.Descripcion
                            })
                            .ToList();
                return new
                {
                    done = true,
                    message = "OK",
                    data = data
                };
            }
        }
        catch (Exception ex)
        {
            log.Error("Excepción Combobox/TipoInforme", ex);
            return new { done = false, message = ex.ToString() };
        }
    }
    private static object GetTipoInforme()
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var data = db.TipoInforme
                            .OrderBy(o => o.Descripcion)
                            .Select(s => new
                            {
                                Value = s.ID,
                                Text = s.Descripcion
                            })
                            .ToList();
                return new
                {
                    done = true,
                    message = "OK",
                    data = data
                };
            }
        }
        catch (Exception ex)
        {
            log.Error("Excepción Combobox/TipoInforme", ex);
            return new { done = false, message = ex.ToString() };
        }
    }
    private static object GetNormas(int norma)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var data = db.Norma
                            .Where(w => w.ID != norma)
                            .OrderBy(o => o.Nombre)
                            .Select(s => new
                            {
                                Value = s.ID,
                                Text = s.Nombre
                            })
                            .ToList();
                return new
                {
                    done = true,
                    message = "OK",
                    data = data
                };
            }
        }
        catch (Exception ex)
        {
            log.Error("Excepción Combobox/Aparatos", ex);
            return new { done = false, message = ex.ToString() };
        }
    }

    private static object GetTipoFuncionamiento()
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var data = db.TipoFuncionamientoAparato
                            .OrderBy(o => o.Descripcion)
                            .Select(s => new
                            {
                                Value = s.ID,
                                Text = s.Descripcion
                            })
                            .ToList();
                return new
                {
                    done = true,
                    message = "OK",
                    data = data
                };
            }
        }
        catch (Exception ex)
        {
            log.Error("Excepción Combobox/Aparatos", ex);
            return new { done = false, message = ex.ToString() };
        }
    }

    private static object GetAparatos()
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var data = db.Aparato
                            .OrderBy(o => o.Nombre)
                            .Select(s => new
                            {
                                Value = s.ID,
                                Text = s.Nombre
                            })
                            .ToList();
                return new
                {
                    done = true,
                    message = "OK",
                    data = data
                };
            }
        }
        catch (Exception ex)
        {
            log.Error("Excepción Combobox/Aparatos", ex);
            return new { done = false, message = ex.ToString() };
        }
    }
    private static object GetInspectores()
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var data = db.Usuario
                            .Where(w => w.UsuarioRol.Where(u => u.Rol == 1).Any())
                            .Where(w => w.Habilitado == true)
                            .Select(s => new
                            {
                                Value = s.NombreUsuario,
                                Text = string.Concat(s.Nombre, " ", s.Apellido)
                            })
                            .ToList();
                return new
                {
                    done = true,
                    message = "OK",
                    data = data
                };
            }
        }
        catch (Exception ex)
        {
            log.Error("Excepción Combobox/TipoNorma", ex);
            return new { done = false, message = ex.ToString() };
        }
    }
    private static object GetTipoNorma()
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var data = db.TipoNorma
                            .Select(s => new
                            {
                                Value = s.ID,
                                Text = s.Descripcion
                            })
                            .ToList();
                return new
                {
                    done = true,
                    message = "OK",
                    data = data
                };
            }
        }
        catch (Exception ex)
        {
            log.Error("Excepción Combobox/TipoNorma", ex);
            return new { done = false, message = ex.ToString() };
        }
    }

    private static object GetClientes()
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var data = db.Cliente
                            .Where(w => w.Habilitado == true)
                            .OrderBy(o => o.Nombre)
                            .Select(s => new
                            {
                                Value = s.ID,
                                Text = s.Nombre
                            })
                            .ToList();
                return new
                {
                    done = true,
                    message = "OK",
                    data = data
                };
            }
        }
        catch (Exception ex)
        {
            log.Error("Excepción Combobox/Clientes", ex);
            return new { done = false, message = ex.ToString() };
        }
    }

    private static object GetHitos(HttpContext post)
    {
        try
        {
            var quotationId = int.Parse(post.Request["quotationId"]);
            using (var db = new CertelEntities())
            {
                var cotizacion = db.Cotizacion.Find(quotationId);
                
                var data = db.TipoHito
                            .Where(w => w.EstadoCotizacionId == cotizacion.EstadoId)
                            .Where(w => cotizacion.EstadoId == 3 && cotizacion.CreadaFase2 == true && w.Id == 19 ? false : true)
                            .Where(w => cotizacion.EstadoId == 3 && cotizacion.CreadaFase2 != true && w.Id == 21 ? false : true)
                            .Select(s => new
                            {
                                Value = s.Id,
                                Text = s.Descripcion
                            })
                            .ToList();
                return new
                {
                    done = true,
                    message = "OK",
                    data = data
                };
            }
        }
        catch (Exception ex)
        {
            log.Error("Excepción Combobox/TitulosNorma", ex);
            return new { done = false, message = ex.ToString() };
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