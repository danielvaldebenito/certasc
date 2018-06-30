<%@ WebHandler Language="C#" Class="DetalleCotizacion" %>

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
using System.Net.Mail;

public class DetalleCotizacion : IHttpHandler
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
            case "getInitialData":
                var quotationId = int.Parse(post.Request["quotationId"]);
                data = GetInitialData(quotationId);
                break;

            case "grid":
                sidx = post.Request["sidx"];
                sord = post.Request["sord"];
                page = int.Parse(post.Request["page"]);
                rows = int.Parse(post.Request["rows"]);
                var Id = int.Parse(post.Request["id"]);
                data = GetHitos(sidx, sord, page, rows, Id);
                break;

            case "quotationItems":
                data = quotationItems(post);
                break;

            case "GetCostosLogisticos":
                data = GetCostosLogisticos(post);
                break;

            case "addHito":
                data = addHito(post);
                break;

            case "sendEmail":
                data = sendEmail(post);
                break;

            case "getmailFormat":
                data = getmailFormat(post);
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

    private static object GetInitialData(int quotationId)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var cotizacion = db.Cotizacion.Find(quotationId);
                if (cotizacion == null) return new { done = false, message = "Error: Cotización no existe" };
                var service = cotizacion.Servicio;
                var inspecciones = service == null ? null : service.Select(s => s.Inspeccion).FirstOrDefault();
                var inspeccionesCount = inspecciones == null ? 0 : inspecciones.Count;
                var ck = "6Fj0rKO5/w2p9/QxOR3NmwO2zWUn4x/zRZGdIhPrDIc=";
                var eck = new Encriptacion(ck, false).newText;
                var quotation = db.Cotizacion
                           .Where(w => w.Id == quotationId)
                           .AsEnumerable()
                           .Select(s => new
                           {
                               //CLIENT DATA
                               clientRut = s.Cliente.Rut,
                               clientName = s.Cliente.Nombre,
                               clientDirection = s.Cliente.Direccion,
                               clientPhone = s.Cliente.TelefonoContacto,
                               clientMail = s.Cliente.EmailContacto,
                               clientGiro = s.Cliente.Giro,
                               clientCiudad = s.Ciudad.Descripcion,
                               clientContactName = s.Cliente.NombreContacto,
                               //QUOTATION DATA
                               quotationFechaDoc = s.FechaDoc == null ? "" : ((DateTime)s.FechaDoc).ToString(("dd-MM-yyyy")),
                               quotationFuenteSolicictud = s.FuenteSolicitud,
                               quotationFormaPago = s.FormaDePago.Descripcion,
                               quotationMoneda = s.Moneda.Descripcion,
                               quotationVendedor = s.Usuario1.Nombre + " " + s.Usuario1.Apellido,
                               quotationValidez = s.FechaValidez == null ? "" : ((DateTime)s.FechaValidez).ToString(("dd-MM-yyyy")),
                               quotationObservation = s.Observacion,
                               quotationNota = s.Nota,
                               quotationStateId = s.EstadoId,
                               quotatationIT = s.IT == null ? 0 : (int)s.IT,
                               quotationEstado = s.EstadoCotizacion.Descripcion,
                               quotationMonedaId = s.MonedaId,
                               ingeniero = inspeccionesCount == 0 ? null : inspecciones.FirstOrDefault().Ingeniero,
                               // Valores Finales
                               eck,
                               quotationValor = s.Valor,
                               quotationDescuento = s.Descuento,
                               quotationRecargo = s.Recargo,
                               quotationTotal = s.Total,
                               //Inspecciones Data
                               inspeccionesCount = inspeccionesCount > 0 ? false : true
                           }).FirstOrDefault();


                return new
                {
                    done = true,
                    data = quotation
                };

            }
        }
        catch (Exception ex)
        {
            return new { done = false, ex = ex.Message };
        }
    }
    private static object GetHitos(string sidx, string sord, int page, int rows, int quotationId)
    {
        try
        {

            using (var db = new CertelEntities())
            {
                var list = db.Hito
                                .Where(w => w.CotizacionId == quotationId);

                int pageSize = rows;
                int totalRecords = list.Count();
                int totalPages = (int)Math.Ceiling((float)totalRecords / (float)pageSize);
                totalPages = totalPages == 0 ? 1 : totalPages;
                var result = list
                    .OrderByDescending(o => o.Id)
                    .ThenBy(o => o.Fecha)
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
                               Fecha = x.Fecha == null ? "" : ((DateTime)x.Fecha).AddHours(5).ToString("dd-MM-yyyy HH:mm:ss", System.Globalization.CultureInfo.CreateSpecificCulture("es-CL")),
                               Hito = x.TipoHito.EstadoCotizacionId == null ? x.TipoHito.Descripcion : x.TipoHito.EstadoCotizacion.Descripcion + " - " + x.TipoHito.Descripcion,
                               Observacion = x.Observacion,
                               FechaCompromiso = x.FechaCompromiso == null ? "" : ((DateTime)x.FechaCompromiso).ToString("dd-MM-yyyy HH:mm:ss"),
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

    private static object addHito(HttpContext post)
    {
        try
        {
            var quotationId = int.Parse(post.Request["quotationId"]);
            var tipoHito = int.Parse(post.Request["tipo"]);
            var observacion = post.Request["observacion"];
            var fechaCompromiso = post.Request["fechaCompromiso"];
            var ingeniero = post.Request["ingeniero"];
            var user = post.Request["user"];

            using (var db = new CertelEntities())
            {
                // Validación Usuario
                var usuario1 = new Encriptacion(user, false).newText;
                var us = db.Usuario.Find(usuario1);
                var from = us.Email;
                var pass = new Encriptacion(us.PassMail, false).newText;
                var name = us.Nombre + " " + us.Apellido;
                if(us.Email == null || us.PassMail == null) {
                    return new { done = false, message = "Configure su cuenta de usuario y contraseña de correo" };
                }
                // Creacion del Hito
                var hito = new Hito
                {
                    CotizacionId = quotationId,
                    Fecha = DateTime.Now,
                    TipoHitoId = tipoHito,
                    Observacion = observacion,
                };

                if (fechaCompromiso != string.Empty)
                    hito.FechaCompromiso = DateTime.ParseExact(fechaCompromiso, "dd-MM-yyyy HH:mm", null);

                db.Hito.Add(hito);

                //Modifica la cotizacion al siguiente estado
                var quotation = db.Cotizacion.Where(w => w.Id == quotationId).FirstOrDefault();

                //Hito de cancelacion de cotizacion
                if (tipoHito == 14 || tipoHito == 15) // Cancela Cotizacion
                {
                    quotation.EstadoId = 4;
                }
                else if (tipoHito == 13) // asigna Ingeniero
                {
                    quotation.EstadoId = 3;
                }


                //Creacion del Servicio 
                if (tipoHito == 13)
                {
                    var servicio = new Servicio
                    {
                        FechaCreacion = DateTime.Now,
                        ClienteID = (int)quotation.ClienteId,
                        IT = quotation.IT.ToString(),
                        EstadoID = 1,
                        CotizacionID = quotationId,
                        fechaCompromiso = (DateTime)hito.FechaCompromiso,
                    };
                    db.Servicio.Add(servicio);


                    var items = db.ItemCotizacion.Where(w => w.CotizacionId == quotationId).ToList();
                    var index = 1;
                    foreach (var i in items)
                    {
                        var noInspection = new[] { 2, 6, 7 };
                        if (noInspection.Contains((int)i.ProductoId)) continue;

                        var uso = db.DestinoProyecto.Where(w => w.Descripcion == i.Uso).FirstOrDefault();

                        for (var h = 0; h < i.Cantidad; h++)
                        {
                            var inspeccion = new Inspeccion
                            {
                                IT = quotation.IT.ToString() + "/" + index.ToString(),
                                AparatoID = i.AparatoID,
                                TipoFuncionamientoID = i.TipoFuncionamientoID,
                                EstadoID = 1,
                                ServicioID = servicio.ID,
                                Ingeniero = ingeniero,
                                FechaCreacion = DateTime.Now,
                                FechaInspeccion = hito.FechaCompromiso,
                                AlturaPisos = i.AlturaPisos,
                                Fase = 1,
                                Ubicacion = i.Ubicacion,
                                Destinatario = i.Cotizacion.Cliente.NombreContacto,
                                NombreProyecto = i.Cotizacion.Cliente.Nombre,
                                NombreEdificio = i.Cotizacion.Cliente.Direccion,


                            };
                            if (uso != null)
                            {
                                inspeccion.DestinoProyectoID = uso.ID;
                            }

                            db.Inspeccion.Add(inspeccion);
                            index += 1;
                        }

                    }

                    var usuario = db.Usuario.Where(w => w.NombreUsuario == ingeniero).FirstOrDefault();


                    if (usuario.Email != null && usuario.Email != "")
                    {
                        SendMailNewService(quotationId, usuario, fechaCompromiso, usuario1);
                    }
                }
                else if(tipoHito == 17) {

                    SendMailChangeDate(quotation.IT.ToString(), fechaCompromiso, ingeniero, usuario1, quotationId, 1);
                }
                else if(tipoHito == 21) {
                    quotation.FechaCompromisoFase2 = DateTime.ParseExact(fechaCompromiso, "dd-MM-yyyy HH:mm", null);
                    SendMailChangeDate(quotation.IT.ToString(), fechaCompromiso, ingeniero, usuario1, quotationId, 2);
                }
                db.SaveChanges();

                return new
                {
                    done = true,
                    message = "Hito creado exitosamente",
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
    private static void SendMailChangeDate(string it, string fechaCompromiso, string ingeniero, string user, int cotizacionId, int fase)
    {
        using (var db = new CertelEntities())
        {

            var service = db.Servicio.FirstOrDefault(s => s.IT == it);
            var usuario = db.Usuario.Find(user);
            var from = usuario.Email;
            var pass = new Encriptacion(usuario.PassMail, false).newText;
            var name = usuario.Nombre + " " + usuario.Apellido;
            if(service != null) {
                var oldDate = service.fechaCompromiso;
                service.fechaCompromiso = DateTime.ParseExact(fechaCompromiso, "dd-MM-yyyy HH:mm", null);
                var inspecciones = service.Inspeccion.Where(w => w.Fase == fase);
                var sameIngeniero = true;
                var oldIngeniero = string.Empty;

                foreach(var i in inspecciones)
                {
                    if (i.Ingeniero != ingeniero) {
                        sameIngeniero = false;
                        oldIngeniero = i.Ingeniero;
                    }

                    i.FechaInspeccion = DateTime.ParseExact(fechaCompromiso, "dd-MM-yyyy HH:mm", null);
                    i.Ingeniero = ingeniero;

                }
                db.SaveChanges();
                var contenido = "";
                var ing = db.Usuario.Find(ingeniero);
                if (sameIngeniero)
                {

                    contenido = string.Format("Estimado {0}: <br> Se cambió la fecha de programación para el servicio <b>IT {1}</b> <br> La fecha es <b>{2}</b>",
                        ing.Nombre + " " + ing.Apellido,
                        it,
                        fechaCompromiso
                    );
                    SendMail(ing.Email, "Cambio de fecha de compromiso", contenido, usuario.Firma, from, pass, name, true);
                }
                else {
                    SendMailReassignService(it, oldIngeniero, oldDate, user);
                    SendMailNewService(cotizacionId, ing, fechaCompromiso, user);
                }

            }
        }

    }
    private static void SendMailReassignService (string it, string old, DateTime? oldDate, string user)
    {
        try
        {
            if (old == string.Empty) return;
            using (var db = new CertelEntities())
            {
                var usuario = db.Usuario.Find(user);
                var from = usuario.Email;
                var pass = new Encriptacion(usuario.PassMail, false).newText;
                var name = usuario.Nombre + " " + usuario.Apellido;
                var oldIngeniero = db.Usuario.Find(old);
                if (oldIngeniero == null) return;
                if (oldIngeniero.Email == null) return;
                var content = string.Format("Estimado {0}: <br> El Servicio IT {1}, agendado para el {2}, fue asignado a otro ingeniero. <br> Muchas gracias.",
                    oldIngeniero.Nombre + " " + oldIngeniero.Apellido,
                    it,
                    ((DateTime)oldDate).ToString("dd-MM-yyyy HH:mm:ss")
                );
                SendMail(oldIngeniero.Email, string.Format("IT {0} Reasignado", it),content, usuario.Firma, from, pass, name, true);
            }
        }
        catch (Exception ex)
        {
            throw ex;
        }
    }
    private static void SendMailNewService(int cotizacionId, Usuario usuario, string fechaCompromiso, string user)
    {
        try
        {

            using (var db = new CertelEntities())
            {
                var usuario1 = db.Usuario.Find(user);
                var from = usuario1.Email;
                var pass = new Encriptacion(usuario1.PassMail, false).newText;
                var name = usuario1.Nombre + " " + usuario1.Apellido;
                var cotizacion = db.Cotizacion.Find(cotizacionId);
                if (cotizacion == null) return;
                var inspecciones = cotizacion
                                    .ItemCotizacion
                                     .GroupBy(g => new
                                     {
                                         descripcionEditada = g.DescripcionEditada,
                                         ubicacion = g.Ubicacion
                                     })
                                     .Select(s => new
                                     {
                                         descripcion = s.Key.descripcionEditada,
                                         ubicacion = s.Key.ubicacion,
                                         cantidad = s.Count()
                                     }).ToList();

                var contenidoInspecciones = "";
                foreach (var i in inspecciones)
                {
                    contenidoInspecciones = contenidoInspecciones + "(" + i.cantidad.ToString() + ")" + " " + i.descripcion + "<br>";
                }

                var ContenidoCorreo = "Estimado {0} <br>" +
                    "Para el dia {1} se le ha asignado un nuevo servicio: IT {3}, a continuacion se indica el detalle de este. <br> <br> " +
                    "{2} <br>" +
                    "Saludos.";

                var CorreoFinal = string.Format(ContenidoCorreo, usuario.Nombre + " " + usuario.Apellido, fechaCompromiso, contenidoInspecciones, cotizacion.IT);

                SendMail(usuario.Email, "Nueva asignacion de Servicio", CorreoFinal, cotizacion.Usuario1.Firma, from, pass, name, true);
            }
        }
        catch(Exception ex)
        {

        }

    }
    private static object quotationItems(HttpContext post)
    {
        try
        {
            var quotationId = int.Parse(post.Request["quotationId"]);
            using (var db = new CertelEntities())
            {

                var data = db.ItemCotizacion.Where(w => w.CotizacionId == quotationId)
                        .AsEnumerable()
                    .Select(s => new
                    {
                        id = s.Id,
                        productoId = s.ProductoId,
                        moneda = s.Cotizacion.MonedaId,
                        cantidad = s.Cantidad,
                        descripcionEditada = s.DescripcionEditada,
                        ubicacion = s.Ubicacion,
                        valorUnitario = s.Unitario,
                        modoDescuento = s.ModoDescuento,
                        //valorDescuento = Math.Round((double)s.ValorDescuento),
                        //precioCliente =  Math.Round((double)((s.Unitario - (s.ValorDescuento == null ? 0 : s.ValorDescuento)) * s.Cantidad))
                        valorDescuento = ((double)s.ValorDescuento),
                        precioCliente = ((double)((s.Unitario - (s.ValorDescuento == null ? 0 : s.ValorDescuento)) * s.Cantidad))


                    })
                    .ToList();

                return new { done = true, data = data };
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


    private static object GetCostosLogisticos(HttpContext post)
    {
        try
        {
            var quotationId = int.Parse(post.Request["quotationId"]);
            using (var db = new CertelEntities())
            {

                var data = db.GastoOperacional.Where(w => w.CotizacionId == quotationId)
                        .AsEnumerable()
                    .Select(s => new
                    {
                        id = s.Id,
                        tipoGasto = s.TipoGasto.Descripcion,
                        valor = s.Valor,
                        cantidad = s.Cantidad,
                        total = s.Cantidad * s.Valor
                    })
                    .ToList();

                return new { done = true, data = data };
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


    private static object sendEmail(HttpContext post)
    {
        try
        {

            var quotationId = int.Parse(post.Request["quotationId"]);
            var destinatarios = post.Request["destinatarios"];
            var contenido = post.Request["contenido"];
            var asunto = post.Request["asunto"];
            var user = post.Request["user"];
            var us = new Encriptacion(user, false).newText;
            using (var db = new CertelEntities())
            {
                var usuario = db.Usuario.Find(us);
                if (usuario == null)
                    return new { done = false, message = "Error, usuario no existe" };
                if (usuario.Email == null || usuario.PassMail == null)
                    return new { done = false, message = "Configure su cuenta de usuario y contraseña de correo" };

                var cotizacion = db.Cotizacion.Where(w => w.Id == quotationId).FirstOrDefault();
                var from = usuario.Email;
                var pass = new Encriptacion(usuario.PassMail, false).newText;
                var name = usuario.Nombre + " " + usuario.Apellido;
                bool enviado = SendMail(destinatarios, asunto, contenido, cotizacion.Usuario1.Firma, from, pass, name, false);

                if (enviado)
                {



                    // Creacion del Hito
                    var hito = new Hito
                    {
                        CotizacionId = quotationId,
                        Fecha = DateTime.Now,
                        TipoHitoId = cotizacion.EstadoId == 1 ? 3 : 5,
                        Observacion = "Email de seguimiento a cliente enviado.",
                    };
                    db.Hito.Add(hito);

                    db.SaveChanges();


                    return new
                    {
                        done = true,
                        message = "Correo Enviado",
                    };
                }
                else
                {
                    return new
                    {
                        done = false,
                        message = "Correo no Enviado, favor validar los datos ingresados",
                    };
                }



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

    private static bool SendMail(string to, string asunto, string correo, string firma, string from, string pass, string name, bool toAgenda)
    {
        try
        {

            string[] mails = to.Split(';');


            var config = new ConfigVariables();
            config.SetData();
            SmtpClient smtpServer = new SmtpClient("mail.certel.cl");
            smtpServer.Port = 25;
            smtpServer.Credentials = new System.Net.NetworkCredential(from, pass);
            MailMessage mail = new MailMessage();
            mail.From = new MailAddress(from, string.Format("Certel - {0}", name));
            //mail.To.Add(to); // para quien
            foreach (string destinatario in mails)
            {
                var isValidMail = IsValidEmail(destinatario.Trim());
                if (isValidMail)
                    mail.To.Add(destinatario.Trim()); // con copia
            };
            if (toAgenda)
                mail.CC.Add(config.agendamiento_mail);
            //mail.CC.Add(config.sender_mail); // con copia
            //mail.CC.Add("agenda@certel.cl");                                 //mail.CC.Add(from);
            mail.Subject = asunto;
            mail.Body = GetHtmlToMail(correo, config.public_logo, firma);
            mail.IsBodyHtml = true;
            mail.BodyEncoding = System.Text.Encoding.GetEncoding("utf-8");
            smtpServer.EnableSsl = true;
            smtpServer.Send(mail);
            return true;
        }
        catch (Exception ex)
        {
            return false;
        }
    }

    private static string GetHtmlToMail(string correo, string logo, string firma)
    {

        var html = "";
        html += "<p>";
        html += correo;
        html += "</p>";
        html += "<hr>";
        html += firma;
        html += "<img src='" + logo + "' width='50%' style='max-width:5cm'>";
        html += "<br>";
        return html;
    }

    private static bool IsValidEmail(string email)
    {
        try
        {
            var addr = new System.Net.Mail.MailAddress(email);
            return addr.Address == email;
        }
        catch
        {
            return false;
        }
    }


    private static object getmailFormat(HttpContext post)
    {
        try
        {
            var quotationId = int.Parse(post.Request["quotationId"]);
            var mail = post.Request["mail"];
            var mailFinal = "";
            using (var db = new CertelEntities())
            {

                var cotizacion = db.Cotizacion
                    .Where(w => w.Id == quotationId)
                    .AsEnumerable()
                    .Select(s => new
                    {
                        fecha = s.FechaCreacion == null ? "" : ((DateTime)s.FechaCreacion).ToString(("dd-MM-yyyy")),
                        nombrecliente = s.Cliente.Nombre
                    })
                    .FirstOrDefault();

                var items = db.ItemCotizacion
                     .Where(w => w.CotizacionId == quotationId)
                     .GroupBy(g => new
                     {
                         descripcionEditada = g.DescripcionEditada,
                         ubicacion = g.Ubicacion
                     })
                     .Select(s => new
                     {
                         descripcion = s.Key.descripcionEditada,
                         ubicacion = s.Key.ubicacion,
                         cantidad = s.Count()
                     }).ToList();

                var contenido = "";
                foreach (var i in items)
                {
                    contenido = contenido + i.cantidad.ToString() + " " + i.descripcion + "; ";
                }


                mailFinal = string.Format(mail, cotizacion.nombrecliente, cotizacion.fecha, contenido);
                return new
                {
                    done = true,
                    mail = mailFinal,
                    data = items
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
            return true;
        }
    }

}