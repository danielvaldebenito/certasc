<%@ WebHandler Language="C#" Class="Inspecciones" %>

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
public class Inspecciones : IHttpHandler
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
        string sidx, sord;
        int page, rows;

        switch (action)
        {
            case "grid":
                sidx = post.Request["sidx"];
                sord = post.Request["sord"];
                page = int.Parse(post.Request["page"]);
                rows = int.Parse(post.Request["rows"]);
                var it = post.Request["it"];
                var desde = post.Request["desde"].Trim();
                DateTime? dDesde = null;
                if (desde != string.Empty)
                    dDesde = DateTime.ParseExact(desde, "dd-MM-yyyy", null);
                var hasta = post.Request["hasta"].Trim();
                DateTime? dHasta = null;
                if (hasta != string.Empty)
                    dHasta = DateTime.ParseExact(hasta, "dd-MM-yyyy", null).AddDays(1);
                string user = post.Request["user"];
                int calificacion = int.Parse(post.Request["calificacion"]);
                data = GetInspecciones(sidx, sord, page, rows, it, dDesde, dHasta, user, calificacion);
                break;
            case "corregirOtF2":
                data = CorregirOtf2(post);
                break;
            case "getInspeccion":
                var id = int.Parse(post.Request["id"]);
                data = GetInspeccion(id);
                break;
            case "getSpecificDataInspeccion":
                var id1 = int.Parse(post.Request["id"]);
                data = GetSpecificDataInspeccion(id1);
                break;
            case "editSpecificDataInspeccion":
                var info = post.Request["info"];
                var inspeccion = int.Parse(post.Request["inspeccion"]);
                var all = bool.Parse(post.Request["all"]);
                data = EditSpecificDataInspeccion(info, inspeccion, all);
                break;
            case "getNormas":
                var inspeccionId1 = int.Parse(post.Request["id"]);
                data = GetNormas(inspeccionId1);
                break;
            case "addOrRemoveNorma":
                var aor_id = int.Parse(post.Request["id"]);
                var aor_state = bool.Parse(post.Request["state"]);
                var aor_inspeccion = int.Parse(post.Request["inspeccion"]);
                data = AddOrRemoveNorma(aor_inspeccion, aor_id, aor_state);
                break;
            case "getCheckList":
                var inspeccionId = int.Parse(post.Request["id"]);
                var titleId = int.Parse(post.Request["title"]);
                data = GetCheckList(inspeccionId, titleId);
                break;
            case "getCheckListF2":
                var inspeccionId2 = int.Parse(post.Request["id"]);
                data = GetCheckListF2(inspeccionId2);
                break;
            case "setCumplimiento":
                var sc_cumplimiento = int.Parse(post.Request["cumplimiento"]);
                var sc_caracteristica = int.Parse(post.Request["caracteristica"]);
                var sc_inspeccion = int.Parse(post.Request["inspeccion"]);
                data = SetCumplimiento(sc_cumplimiento, sc_caracteristica, sc_inspeccion);
                break;
            case "writeObservacion":
                data = WriteObservacion(post);
                break;
            case "uploadImage":
                data = UploadImage(post);
                break;
            case "uploadImageTecnica":
                data = UploadImageTecnica(post);
                break;
            case "getPhotos":
                data = GetPhotos(post);
                break;
            case "removePhoto":
                data = RemovePhoto(post);
                break;
            case "saveCalificacion":
                data = SaveCalificacion(post);
                break;
            case "saveObservacionTecnica":
                data = SaveObservacionTecnica(post);
                break;
            case "editOT":
                data = EditOT(post);
                break;
            case "getObservacionesTecnicas":
                data = GetObservacionesTecnicas(post);
                break;
            case "getObservacionesTecnicasF2":
                data = GetObservacionesTecnicasF2(post);
                break;
            case "removeObservacionTecnica":
                data = RemoveObservacionTecnica(post);
                break;
            case "setTipoInforme":
                data = SetTipoInforme(post);
                break;
            case "deleteInspeccion":
                data = DeleteInspeccion(post);
                break;
            case "soloAprobar":
                data = SoloAprobar(post);
                break;
            case "copy":
                data = Copy(post);
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
            log.Error("Inspecciones/Plantilla", ex);
            return new { done = false, message = ex.ToString() };
        }
    }
    private static object Copy(HttpContext post)
    {
        try
        {
            var array = post.Request["array"];
            var toNew = bool.Parse(post.Request["toNew"]);
            var from = int.Parse(post.Request["from"]);
            var to = int.Parse(post.Request["to"]);
            var deserializer = new JavaScriptSerializer();
            var list = deserializer.Deserialize<List<int>>(array);
            using (var db = new CertelEntities())
            {
                var inspToCopy = db.Inspeccion.Find(from);
                if (inspToCopy == null) return new { done = false, message = "Error: Inspección no existe" };

                Inspeccion inspToPaste;
                if (toNew)
                    inspToPaste = new Inspeccion();
                else
                    inspToPaste = db.Inspeccion.Find(to);
                var service = inspToCopy.Servicio;
                if (list.Contains(1)) // Datos Generales
                {
                    inspToPaste.AlturaPisos = inspToCopy.AlturaPisos;
                    inspToPaste.AparatoID = inspToCopy.AparatoID;
                    inspToPaste.Destinatario = inspToCopy.Destinatario;
                    inspToPaste.DestinoProyectoID = inspToCopy.DestinoProyectoID;
                    inspToPaste.Fase = inspToCopy.Fase;
                    inspToPaste.FechaInspeccion = inspToCopy.FechaInspeccion;
                    inspToPaste.FechaInstalacion = inspToCopy.FechaInstalacion;
                    inspToPaste.Ingeniero = inspToCopy.Ingeniero;
                    inspToPaste.NombreEdificio = inspToCopy.NombreEdificio;
                    inspToPaste.NombreProyecto = inspToCopy.NombreProyecto;
                    inspToPaste.PermisoEdificacion = inspToCopy.PermisoEdificacion;
                    inspToPaste.RecepcionMunicipal = inspToCopy.RecepcionMunicipal;
                    inspToPaste.ServicioID = inspToCopy.ServicioID;
                    inspToPaste.TipoFuncionamientoID = inspToCopy.TipoFuncionamientoID;
                    inspToPaste.Ubicacion = inspToCopy.Ubicacion;

                    inspToPaste.EstadoID = 1;
                    inspToPaste.FechaCreacion = DateTime.Now;
                    inspToPaste.IT = toNew ? GetNewIt(service) : inspToPaste.IT;
                }
                if (toNew)
                    db.Inspeccion.Add(inspToPaste);
                db.SaveChanges();
                var inspeccionId = inspToPaste.ID;
                if (list.Contains(2) && inspToCopy.Fase == 1) // Datos específicos (solo para fase 1)
                {
                    // Remove exists
                    var especificosOld = inspToPaste.ValoresEspecificos.ToList();
                    foreach (var e in especificosOld)
                    {
                        db.ValoresEspecificos.Remove(e);
                    }
                    db.SaveChanges();
                    // new
                    var especificos = inspToCopy.ValoresEspecificos;

                    foreach (var e in especificos)
                    {
                        var especifico = new ValoresEspecificos
                        {
                            EspecificoID = e.EspecificoID,
                            InspeccionID = inspeccionId,
                            Valor = e.Valor
                        };
                        db.ValoresEspecificos.Add(especifico);

                    }
                }
                if (list.Contains(3)) // Normas y tipo de informe
                {
                    // Tipo informe
                    inspToPaste.TipoInformeID = inspToCopy.TipoInformeID;
                    // normas
                    // remove exist
                    var inspNormasOld = inspToPaste.InspeccionNorma.ToList();
                    foreach (var ni in inspNormasOld)
                    {
                        db.InspeccionNorma.Remove(ni);
                    }
                    db.SaveChanges();
                    // new
                    var inspNormas = inspToCopy.InspeccionNorma;
                    foreach (var ni in inspNormas)
                    {
                        var inspNorma = new InspeccionNorma
                        {
                            InspeccionID = inspeccionId,
                            NormaID = ni.NormaID
                        };
                        db.InspeccionNorma.Add(inspNorma);
                    }
                }
                if (list.Contains(4)) // Checklist y calificación
                {
                    // Calificacion
                    inspToPaste.Calificacion = inspToCopy.Calificacion;

                    // Cumplimiento
                    // remove old
                    var cumplimientoOld = inspToPaste.Cumplimiento.ToList();
                    foreach (var cum in cumplimientoOld)
                    {
                        if (list.Contains(5) && inspToCopy.Fase == 1)
                        {
                            var photosOld = cum.Fotografias.ToList();
                            foreach(var p in photosOld)
                            {
                                db.Fotografias.Remove(p);
                            }
                            db.SaveChanges();
                        }
                        db.Cumplimiento.Remove(cum);
                    }
                    db.SaveChanges();
                    // new
                    var cumplimiento = inspToCopy.Cumplimiento;
                    foreach (var cum in cumplimiento)
                    {
                        var cumplim = new Cumplimiento
                        {
                            CaracteristicaID = cum.CaracteristicaID,
                            EvaluacionID = cum.EvaluacionID,
                            InspeccionID = inspeccionId,
                            Fecha = DateTime.Now
                        };
                        db.Cumplimiento.Add(cumplim);
                        db.SaveChanges();
                        if (list.Contains(5) && inspToCopy.Fase == 1) // Observaciones y fotografias normativas
                        {
                            cumplim.Observacion = cum.Observacion;
                            var photos = cum.Fotografias;
                            foreach (var p in photos)
                            {
                                var photografy = new Fotografias
                                {
                                    CumplimientoID = cumplim.ID,
                                    URL = p.URL
                                };
                                db.Fotografias.Add(photografy);

                            }
                        }
                    }
                }
                if (list.Contains(6) && inspToCopy.Fase == 1)
                {
                    // remove
                    var observacionesTecnicasOld = inspToPaste.ObservacionTecnica.ToList();
                    foreach(var ot in observacionesTecnicasOld)
                    {
                        var photoTec = ot.FotografiaTecnica.ToList();
                        foreach(var p in photoTec)
                        {
                            db.FotografiaTecnica.Remove(p);
                        }
                        db.SaveChanges();
                        db.ObservacionTecnica.Remove(ot);
                    }
                    db.SaveChanges();
                    var observacionesTecnicas = inspToCopy.ObservacionTecnica;
                    foreach (var ot in observacionesTecnicas)
                    {
                        var obsTec = new ObservacionTecnica
                        {
                            CorregidoEnFase2 = false,
                            InspeccionID = inspeccionId,
                            Texto = ot.Texto
                        };
                        db.ObservacionTecnica.Add(obsTec);
                        db.SaveChanges();
                        var photos = ot.FotografiaTecnica;
                        foreach (var p in photos)
                        {
                            var photoTec = new FotografiaTecnica
                            {
                                ObservacionTecnicaID = obsTec.ID,
                                URL = p.URL
                            };
                            db.FotografiaTecnica.Add(photoTec);

                        }
                    }
                }
                db.SaveChanges();
                return new { done = true, message = "La inspección ha sido copiada al IT " + inspToPaste.IT };
            }
        }
        catch (Exception ex)
        {
            log.Error("Inspecciones/Copy", ex);
            return new { done = false, message = ex.ToString() };
        }
    }
    private static string GetNewIt(Servicio service)
    {
        var inspecciones = service.Inspeccion;
        var digit = 0;
        while(true) {
            digit++;
            var existsInspeccion = inspecciones.Any(a => int.Parse(a.IT.Split('/')[1]) == digit);
            if (!existsInspeccion) break;
        }
        return string.Format("{0}/{1}", service.IT, digit);


    }
    private static object SoloAprobar(HttpContext post)
    {
        try
        {
            var usuario = post.Request["usuario"];

            var id = int.Parse(post.Request["id"]);
            using (var db = new CertelEntities())
            {
                var inspeccion = db.Inspeccion.Find(id);
                if (inspeccion == null) return new { done = false, message = "Error: Inspección no existe" };

                var user = new Encriptacion(usuario, false).newText;
                inspeccion.Aprobador = user;
                inspeccion.FechaAprobacion = DateTime.Now;
                db.SaveChanges();
                return new { done = true, message = "La inspección ha sido aprobada" };
            }
        }
        catch (Exception ex)
        {
            log.Error("Inspecciones/Aprobar", ex);
            return new { done = false, message = ex.ToString() };
        }
    }
    private static object CorregirOtf2(HttpContext post)
    {
        try
        {
            var id = int.Parse(post.Request["id"]);
            var ok = bool.Parse(post.Request["ok"]);
            using (var db = new CertelEntities())
            {
                var ot = db.ObservacionTecnica.Find(id);
                ot.CorregidoEnFase2 = ok;
                db.SaveChanges();
                return new { done = true, message = "OK", ok = ok };
            }
        }
        catch (Exception ex)
        {
            log.Error("Inspecciones/Plantilla", ex);
            return new { done = false, message = ex.ToString() };
        }
    }
    private static object EditOT(HttpContext post)
    {
        try
        {
            var id = int.Parse(post.Request["id"]);
            var texto = post.Request["text"];
            using (var db = new CertelEntities())
            {
                var ot = db.ObservacionTecnica.Find(id);
                ot.Texto = texto;
                db.SaveChanges();
                return new { done = true, message = "Registro Actualizado" };
            }
        }
        catch (Exception ex)
        {
            log.Error("Inspecciones/EditOT", ex);
            return new { done = false, message = ex.ToString() };
        }
    }
    private static object DeleteInspeccion(HttpContext post)
    {
        try
        {
            var inspeccionId = int.Parse(post.Request["id"]);
            using (var db = new CertelEntities())
            {
                var inspeccion = db.Inspeccion.Find(inspeccionId);
                if (inspeccion == null)
                    return new { done = false, message = "Esta inspección ya no existe" };
                var path = HttpContext.Current.Server.MapPath("~/fotos/");

                if (inspeccion.Inspeccion1.Any())
                    return new { done = false, message = "Esta inspección tiene una inspección en Fase 2, cuando finalice la fase 2, puede eliminarla." };
                var cumplimientos = inspeccion.Cumplimiento.ToList();
                cumplimientos.ForEach(f =>
                {
                    var fotos = f.Fotografias.ToList();
                    fotos.ForEach(ff =>
                    {
                        if (File.Exists(path + ff.URL))
                            File.Delete(path + ff.URL);
                        db.Fotografias.Remove(ff);
                    });
                    db.SaveChanges();
                    db.Cumplimiento.Remove(f);
                });
                var obsTecns = inspeccion.ObservacionTecnica.ToList();
                obsTecns.ForEach(f =>
                {
                    var fotos = f.FotografiaTecnica.ToList();
                    fotos.ForEach(ff =>
                    {
                        if (File.Exists(path + ff.URL))
                            File.Delete(path + ff.URL);
                        db.FotografiaTecnica.Remove(ff);
                    });
                    db.SaveChanges();
                    db.ObservacionTecnica.Remove(f);
                });
                var inspeccionNorma = inspeccion.InspeccionNorma.ToList();
                inspeccionNorma.ForEach(f => { db.InspeccionNorma.Remove(f); });
                var informes = inspeccion.Informe.ToList();
                informes.ForEach(f => { db.Informe.Remove(f); });
                var valores = inspeccion.ValoresEspecificos.ToList();
                valores.ForEach(f => { db.ValoresEspecificos.Remove(f); });
                db.SaveChanges();
                db.Inspeccion.Remove(inspeccion);

                db.SaveChanges();
                return new { done = true, message = "Inspección Eliminada" };

            }
        }
        catch (Exception ex)
        {
            log.Error("Inspecciones/DeleteInspeccion", ex);
            return new { done = false, message = ex.ToString() };
        }
    }
    private static object SetTipoInforme(HttpContext post)
    {
        try
        {
            var inspeccionId = int.Parse(post.Request["inspeccion"]);
            var informeId = int.Parse(post.Request["informe"]);
            using (var db = new CertelEntities())
            {
                var inspeccion = db.Inspeccion
                                    .Find(inspeccionId);
                inspeccion.TipoInformeID = informeId;
                var if2 = inspeccion.Inspeccion1;
                foreach (var i in if2)
                {
                    i.TipoInformeID = informeId;
                }
                db.SaveChanges();
                return new { done = true, message = "OK" };
            }
        }
        catch (Exception ex)
        {
            log.Error("Inspecciones/SetTipoInforme", ex);
            return new { done = false, message = ex.ToString() };
        }
    }
    private static object RemoveObservacionTecnica(HttpContext post)
    {
        try
        {
            var id = int.Parse(post.Request["id"]);
            using (var db = new CertelEntities())
            {

                var observacion = db.ObservacionTecnica
                                    .Find(id);
                if (observacion == null) return new { done = false, message = "Ha ocurrido un error" };

                var fotos = db.FotografiaTecnica
                                .Where(w => w.ObservacionTecnicaID == id)
                                .ToList();
                foreach (var f in fotos)
                {
                    var url = f.URL;
                    var others = db.FotografiaTecnica
                                    .Where(w => w.URL == url)
                                    .Count();
                    var fileName = HttpContext.Current.Server.MapPath("~/fotos/") + f.URL;

                    db.FotografiaTecnica.Remove(f);
                    db.SaveChanges();

                    if(others <= 1)
                    {
                        if(File.Exists(fileName))
                            File.Delete(fileName);
                    }
                }

                db.ObservacionTecnica.Remove(observacion);
                db.SaveChanges();
                return new { done = true, message = "Observación eliminada exitosamente" };
            }
        }
        catch (Exception ex)
        {
            log.Error("Inspecciones/RemoveObservacionTecnica", ex);
            return new { done = false, message = ex.ToString() };
        }
    }
    private static object SaveObservacionTecnica(HttpContext post)
    {
        try
        {
            var inspeccionId = int.Parse(post.Request["inspeccionId"]);
            var texto = post.Request["texto"];
            using (var db = new CertelEntities())
            {
                var observacion = new ObservacionTecnica
                {
                    InspeccionID = inspeccionId,
                    Texto = texto,
                };
                db.ObservacionTecnica.Add(observacion);
                db.SaveChanges();
                return new { done = true, message = "ok", id = observacion.ID };
            }
        }
        catch (Exception ex)
        {
            log.Error("Inspecciones/SaveObservacionTecnica", ex);
            return new { done = false, message = ex.ToString() };
        }
    }
    private static object GetObservacionesTecnicas(HttpContext post)
    {
        try
        {
            var inspeccionId = int.Parse(post.Request["inspeccionId"]);
            using (var db = new CertelEntities())
            {
                var observaciones = db.ObservacionTecnica
                                        .Where(w => w.InspeccionID == inspeccionId)
                                        .Select(s => new
                                        {
                                            Id = s.ID,
                                            Texto = s.Texto,
                                            Image = s.FotografiaTecnica.Select(ss => "fotos/" + ss.URL).FirstOrDefault() ?? string.Empty
                                        })
                                        .ToList();
                return observaciones;
            }
        }
        catch (Exception ex)
        {
            log.Error("Inspecciones/GetObservacionesTecnicas", ex);
            return new { done = false, message = ex.ToString() };
        }
    }

    private static object GetObservacionesTecnicasF2(HttpContext post)
    {
        try
        {
            var inspeccionId = int.Parse(post.Request["inspeccionId"]);
            using (var db = new CertelEntities())
            {
                var inspeccion = db.Inspeccion.Find(inspeccionId);
                var inspeccionF1 = inspeccion.InspeccionFase1;
                var observaciones = db.ObservacionTecnica
                                        .Where(w => w.InspeccionID == inspeccionF1)
                                        .Select(s => new
                                        {
                                            Id = s.ID,
                                            Texto = s.Texto,
                                            Image = s.FotografiaTecnica.Select(ss => "fotos/" + ss.URL).FirstOrDefault() ?? string.Empty,
                                            Corregido = s.CorregidoEnFase2
                                        })
                                        .ToList();
                return observaciones;
            }
        }
        catch (Exception ex)
        {
            log.Error("Inspecciones/GetObservacionesTecnicasF2", ex);
            return new { done = false, message = ex.ToString() };
        }
    }


    private static object SaveCalificacion(HttpContext post)
    {
        try
        {

            var inspeccionId = int.Parse(post.Request["inspeccionId"]);
            var val = int.Parse(post.Request["val"]);
            var dias = int.Parse(post.Request["dias"]);
            var creafase2 = bool.Parse(post.Request["creafase2"]);
            var mensaje = "Guardado";
            using (var db = new CertelEntities())
            {
                var inspeccion = db.Inspeccion
                                    .Find(inspeccionId);
                inspeccion.Calificacion = val;
                if (val == 0 || val == 2)
                    inspeccion.DiasPlazo = dias;
                db.SaveChanges();
                bool crea;
                if (creafase2)
                {
                    if (inspeccion.Fase == 1)
                    {
                        if (!inspeccion.Inspeccion1.Any())
                        {
                            var fase2 = new Inspeccion
                            {
                                EstadoID = 1,
                                Fase = 2,
                                FechaCreacion = DateTime.Now,
                                InspeccionFase1 = inspeccion.ID,
                                AparatoID = inspeccion.AparatoID,
                                AlturaPisos = inspeccion.AlturaPisos,
                                Destinatario = inspeccion.Destinatario,
                                DestinoProyectoID = inspeccion.DestinoProyectoID,
                                FechaInstalacion = inspeccion.FechaInstalacion,
                                Ingeniero = inspeccion.Ingeniero,
                                IT = inspeccion.IT,
                                NombreEdificio = inspeccion.NombreEdificio,
                                NombreProyecto = inspeccion.NombreProyecto,
                                TipoInformeID = inspeccion.TipoInformeID,
                                TipoFuncionamientoID = inspeccion.TipoFuncionamientoID,
                                ServicioID = inspeccion.ServicioID,
                                Numero = inspeccion.Numero,
                                Ubicacion = inspeccion.Ubicacion,
                                PermisoEdificacion = inspeccion.PermisoEdificacion,
                                RecepcionMunicipal = inspeccion.RecepcionMunicipal,


                            };

                            if (inspeccion.FechaEntrega.HasValue)
                                fase2.FechaInspeccion = inspeccion.FechaEntrega.Value.AddDays(dias);

                            db.Inspeccion.Add(fase2);
                            db.SaveChanges();


                            var inspeccionNormaOriginal = inspeccion.InspeccionNorma.ToList();
                            foreach (var i in inspeccionNormaOriginal)
                            {
                                var inspeccionNorma = new InspeccionNorma
                                {
                                    InspeccionID = fase2.ID,
                                    NormaID = i.NormaID
                                };
                                db.InspeccionNorma.Add(inspeccionNorma);
                                db.SaveChanges();
                            }

                            mensaje = "La calificación fue guardada. La fase 2 ha sido creada";

                        }
                        else
                        {
                            mensaje = "La calificación fue guardada, pero no se creó una Fase 2, porque ya existe una.";
                        }
                    }
                    else
                    {
                        mensaje = "La calificación ha sido guardada. Se especificará en el informe, que se creará una fase 3, materia de otra cotización";
                    }
                    crea = true;
                }
                else
                {
                    crea = false;
                }
                inspeccion.CreaFaseSiguiente = crea;
                db.SaveChanges();
                return new { done = true, message = mensaje };
            }

        }

        catch (Exception ex)
        {
            log.Error("Inspecciones/Plantilla", ex);
            return new { done = false, message = ex.ToString() };
        }
    }
    private static object RemovePhoto(HttpContext post)
    {
        try
        {
            var id = int.Parse(post.Request["id"]);
            using (var db = new CertelEntities())
            {
                var photo = db.Fotografias
                                .Where(w => w.ID == id)
                                .FirstOrDefault();

                if (photo == null)
                {
                    return new { done = false, message = "Fotografía ya no existe" };
                }
                var url = photo.URL;
                var moreThanOne = db.Fotografias
                                    .Where(w => w.URL == url)
                                    .Count();
                if(moreThanOne <= 1)
                {
                    var fileName = HttpContext.Current.Server.MapPath("~/fotos/") + photo.URL;
                    if(File.Exists(fileName))
                        File.Delete(fileName);
                }

                db.Fotografias.Remove(photo);
                db.SaveChanges();

                return new { done = true, message = "Fotografía borrada exitosamente" };
            }
        }
        catch (Exception ex)
        {
            log.Error("Inspecciones/Plantilla", ex);
            return new { done = false, message = ex.ToString() };
        }
    }

    private static object GetPhotos(HttpContext post)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var car = int.Parse(post.Request["caracteristica"]);
                var insp = int.Parse(post.Request["inspeccion"]);
                var cumplimiento = db.Cumplimiento
                                    .Where(w => w.InspeccionID == insp)
                                    .Where(w => w.CaracteristicaID == car)
                                    .FirstOrDefault();
                if (cumplimiento == null)
                    return new { done = false, code = 2, message = "Primero seleccione una evaluación" };

                var photos = db.Fotografias
                                .Where(w => w.CumplimientoID == cumplimiento.ID)
                                .ToList();

                if (photos.Count() == 0)
                    return new { done = false, code = 1, message = "No hay fotografías" };

                return new
                {
                    done = true,
                    message = "No hay fotografías",
                    photos = photos
                            .Select(s => new
                            {
                                Id = s.ID,
                                URL = "fotos/" + s.URL
                            })
                };

            }
        }
        catch (Exception ex)
        {
            log.Error("Inspecciones/Plantilla", ex);
            return new { done = false, code = 0, message = ex.ToString() };
        }
    }

    private static object UploadImage(HttpContext post)
    {
        try
        {

            var inspeccion = int.Parse(post.Request["inspeccion"]);
            var caracteristica = int.Parse(post.Request["caracteristica"]);
            string str_image = string.Empty;
            var basePath = HttpContext.Current.Server.MapPath("~/fotos/");
            using (var db = new CertelEntities())
            {
                var cumplimiento = db.Cumplimiento
                                    .Where(w => w.CaracteristicaID == caracteristica)
                                    .Where(w => w.InspeccionID == inspeccion)
                                    .FirstOrDefault();
                if (cumplimiento == null)
                    return new { done = false, message = "Primero ingrese una evaluación para esta característica" };

                HttpPostedFile file = post.Request.Files[0];
                string fileName = file.FileName;
                string fileExtension = file.ContentType;
                if (string.IsNullOrEmpty(fileName)) return new { done = false, message = "Error: No se ha recibido correctamente el archivo" };

                var date = DateTime.Now.ToString("ddMMyyyyHHmmss");
                fileExtension = Path.GetExtension(fileName);
                str_image = inspeccion + "_" + caracteristica + "_" + date + fileExtension;
                string pathToTempSave = basePath + "Temp_" + str_image;
                file.SaveAs(pathToTempSave);
                var destinyPath = basePath + str_image;
                // Save to DB
                var photo = new Fotografias
                {
                    CumplimientoID = cumplimiento.ID,
                    URL = str_image,
                };
                db.Fotografias.Add(photo);
                db.SaveChanges();

                // Resize
                if (ResizeImage(pathToTempSave, destinyPath))
                {
                    return new { done = true, message = "Fotografía agregada correctamente" };
                }
                else {
                    file.SaveAs(basePath + str_image);
                    return new { done = true, message = "Fotografía agregada, pero no fue posible reducir la imagen" };
                }
            }
        }
        catch (Exception ex)
        {
            log.Error("Inspecciones/UploadImage", ex);
            return new { done = false, message = ex.ToString() };
        }
    }
    public static bool ResizeImage(string temppath, string destinyPath)
    {
        try
        {
            var maxWidth = 500;
            using (var i = Image.FromFile(temppath))
            {
                var width = i.Width;
                var height = i.Height;
                var maxHeight = (int)(height * maxWidth / width);
                Image resized = new Bitmap(i, new Size(maxWidth, maxHeight));
                resized.Save(destinyPath);
            }
            Thread.Sleep(100);
            if (File.Exists(temppath))
                File.Delete(temppath);
            return true;
        }
        catch(Exception ex)
        {
            return false;
        }

    }
    private static object UploadImageTecnica(HttpContext post)
    {
        try
        {

            var obs = post.Request["obs"];
            var inspeccionId = int.Parse(post.Request["inspeccion"]);
            string str_image = string.Empty;
            var basePath = HttpContext.Current.Server.MapPath("~/fotos/");
            using (var db = new CertelEntities())
            {

                HttpPostedFile file = post.Request.Files[0];
                string fileName = file.FileName;
                string fileExtension = file.ContentType;
                if (string.IsNullOrEmpty(fileName)) return new { done = false, message = "Error: No se ha recibido correctamente el archivo" };

                var date = DateTime.Now.ToString("ddMMyyyyHHmmss");
                fileExtension = Path.GetExtension(fileName);
                str_image = inspeccionId + "_OT_" + date + fileExtension;
                string pathToTempSave = basePath  + "Temp_" + str_image;
                file.SaveAs(pathToTempSave);
                var destinyPath = basePath + str_image;


                var observacion = new ObservacionTecnica
                {
                    InspeccionID = inspeccionId,
                    Texto = obs,
                };
                db.ObservacionTecnica.Add(observacion);
                db.SaveChanges();
                // Save To BBDD
                var photo = new FotografiaTecnica
                {
                    ObservacionTecnicaID = observacion.ID,
                    URL = str_image,

                };
                db.FotografiaTecnica.Add(photo);
                db.SaveChanges();
                if(ResizeImage(pathToTempSave, destinyPath))
                    return new { done = true, message = "Observación y Fotografía agregadas correctamente" };
                else
                {
                    file.SaveAs(basePath + str_image);
                    return new { done = true, message = "Observación y agregada correctamente. La fotografía también se agregó pero no se pudo reducir el tamaño." };
                }

            }


        }
        catch (Exception ex)
        {
            log.Error("Inspecciones/UploadImageTecnica", ex);
            return new { done = false, message = ex.ToString() };
        }
    }
    private static object WriteObservacion(HttpContext post)
    {
        try
        {
            var observacion = post.Request["observacion"];
            var inspeccion = int.Parse(post.Request["inspeccion"]);
            var caracteristica = int.Parse(post.Request["caracteristica"]);

            using (var db = new CertelEntities())
            {
                var exists = db.Cumplimiento
                                .Where(w => w.InspeccionID == inspeccion)
                                .Where(w => w.CaracteristicaID == caracteristica)
                                .FirstOrDefault();
                if (exists == null)
                {
                    return new
                    {
                        done = false,
                        message = "Primero ingrese un cumplimiento"
                    };
                }
                else
                {
                    exists.Observacion = observacion;
                    exists.Fecha = DateTime.Now;
                }
                db.SaveChanges();
                return new
                {
                    done = true,
                    message = "Observación guardada"
                };
            }
        }
        catch (Exception ex)
        {
            log.Error("Inspecciones/Plantilla", ex);
            return new { done = false, message = ex.ToString() };
        }
    }

    private static object SetCumplimiento(int eval, int caracteristica, int inspeccion)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var exists = db.Cumplimiento
                                .Where(w => w.CaracteristicaID == caracteristica)
                                .Where(w => w.InspeccionID == inspeccion)
                                .FirstOrDefault();
                if (exists == null)
                {
                    var evaluacion = new Cumplimiento
                    {
                        CaracteristicaID = caracteristica,
                        EvaluacionID = eval,
                        InspeccionID = inspeccion,
                        Fecha = DateTime.Now
                    };
                    db.Cumplimiento.Add(evaluacion);

                }
                else
                {
                    exists.EvaluacionID = eval;
                    exists.Fecha = DateTime.Now;

                }


                db.SaveChanges();
                return new
                {
                    done = true,
                    message = "Cumplimiento Guardado oK"
                };
            }
        }
        catch (Exception ex)
        {
            log.Error("Inspecciones/SetCumplimiento", ex);
            return new { done = false, message = ex.ToString() };
        }
    }


    private static object AddOrRemoveNorma(int inspeccionId, int id, bool check)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var inspeccion = db.Inspeccion.Find(inspeccionId);
                var inspeccionesEnFase2 = inspeccion.Inspeccion1;
                if (check)
                {
                    var any = db.InspeccionNorma.Any(a => a.InspeccionID == inspeccionId && a.NormaID == id);
                    if (!any)
                    {
                        var ni = new InspeccionNorma
                        {
                            InspeccionID = inspeccionId,
                            NormaID = id
                        };
                        db.InspeccionNorma.Add(ni);
                        foreach (var if2 in inspeccionesEnFase2)
                        {
                            var any1 = db.InspeccionNorma.Any(a => a.InspeccionID == if2.ID && a.NormaID == id);
                            if (!any1)
                            {
                                var ni2 = new InspeccionNorma
                                {
                                    InspeccionID = if2.ID,
                                    NormaID = id
                                };
                                db.InspeccionNorma.Add(ni2);
                            }
                        }
                        db.SaveChanges();
                    }

                    //var norma = db.Norma.Find(id);
                    //var asociadas = norma.NormasAsociadas.ToList();
                    //foreach(var a in asociadas)
                    //{
                    //    var ni1 = new InspeccionNorma
                    //    {
                    //        InspeccionID = inspeccion,
                    //        NormaID = a.Norma1.ID
                    //    };
                    //    db.InspeccionNorma.Add(ni1);
                    //    db.SaveChanges();
                    //}

                    return new
                    {
                        done = true,
                        message = "OK"
                    };
                }
                else
                {
                    var ids = inspeccionesEnFase2.Select(s => (int?)s.ID);
                    var ni = db.InspeccionNorma
                                .Where(w => w.NormaID == id)
                                .Where(w => w.InspeccionID == inspeccionId);
                    if (!ni.Any())
                        return new
                        {
                            done = false,
                            message = "Ha ocurrido un error"
                        };
                    foreach (var f in ni)
                    {
                        db.InspeccionNorma.Remove(f);
                    }

                    foreach (var if2 in inspeccionesEnFase2)
                    {
                        var normainsp = db.InspeccionNorma
                                        .Where(w => w.NormaID == id)
                                        .Where(w => w.InspeccionID == if2.ID)
                                        .ToList();
                        normainsp.ForEach(f => { db.InspeccionNorma.Remove(f); });

                    }
                    db.SaveChanges();
                    return new
                    {
                        done = true,
                        message = "OK"
                    };
                }

            }
        }
        catch (Exception ex)
        {
            log.Error("Inspecciones/AddOrRemoveNorma", ex);
            return new { done = false, message = ex.ToString() };
        }
    }
    private static object GetNormas(int inspeccion)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var inspec = db.Inspeccion
                                .Find(inspeccion);

                var ni = inspec.InspeccionNorma
                            .Select(s => s.NormaID)
                            .ToList();

                var normas = db.Norma
                                .Where(w => w.NormasAsociadas.Any() || w.Principal == true)
                                .Select(s => new
                                {
                                    Id = s.ID,
                                    Nombre = s.Nombre,
                                    Checked = ni.Contains(s.ID),
                                    Hijas = s.NormasAsociadas.Select(ss => ss.ID).ToList()
                                })
                                .ToList();
                var informes = db.TipoInforme
                                    .Select(s => new
                                    {
                                        Id = s.ID,
                                        Nombre = s.Descripcion,
                                        Cheched = inspec.TipoInformeID == s.ID
                                    }).ToList();
                return new { done = true, normas = normas, informes = informes };
            }
        }
        catch (Exception ex)
        {
            log.Error("Inspecciones/GetNormas", ex);
            return new { done = false, message = ex.ToString() };
        }
    }

    private static object GetCheckList(int inspeccionId, int title)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var inspeccion = db.Inspeccion.Find(inspeccionId);
                var titles = new TitulosInspeccion(inspeccion).ListaInspeccion;
                var first = titles.First();
                List<Norma> normasList = new List<Norma>();
                var normasReguladoras = db.InspeccionNorma
                                            .Where(w => w.InspeccionID == inspeccionId)
                                            .Where(w => w.Norma.Principal == true)
                                            ;

                foreach (var n in normasReguladoras)
                {
                    normasList.Add(n.Norma);
                    var asociadas = n.Norma.NormasAsociadas;
                    foreach (var a in asociadas)
                    {
                        normasList.Add(a.Norma1);
                    }
                }
                if (normasList.Count == 0)
                {
                    return new { done = false, code = 1, message = "No hay normas asignadas a esta inspección. ¿Desea agregarlas?" };
                }
                var data =
                    normasList
                        .Where(w => title == 0 ? true : w.Titulo.Any(a => a.ID == title))
                        .Select(n => new
                        {
                            Id = n.ID,
                            Text = n.Nombre,
                            Titulos = n.Titulo
                                    .Where(w => title == 0 ? true : w.ID == title)
                                    .Select(s => new
                                    {
                                        Id = s.ID,
                                        Text = s.Texto,
                                        Requisitos = s.Requisito
                                                .Where(w => w.Habilitado == true)
                                                  .Select(r => new
                                                  {
                                                      Id = r.ID,
                                                      Text = r.Descripcion,
                                                      Caracteristicas = r.Caracteristica
                                                                            .Where(w => w.Habilitado == true)
                                                                            .Select(c => new
                                                                            {
                                                                                Id = c.ID,
                                                                                Text = c.Descripcion,
                                                                                Cumplimiento = c.Cumplimiento
                                                                                                .Where(w => w.InspeccionID == inspeccionId)
                                                                                                .Select(cu => new
                                                                                                {
                                                                                                    Id = cu.EvaluacionID,
                                                                                                    Text = cu.Evaluacion.Descripcion,
                                                                                                    Observacion = cu.Observacion,
                                                                                                    HasFotos = cu.Fotografias.Any()
                                                                                                })
                                                                                                .FirstOrDefault()
                                                                            })
                                                                            .ToList()
                                                  })
                                                  .ToList()
                                    })
                            .ToList()
                        })
                        .ToList();

                return new { done = true, data = data };
            }
        }
        catch (Exception ex)
        {
            log.Error("Inspecciones/GetCheckList", ex);
            return new { done = false, code = 0, message = ex.ToString() };
        }
    }
    private static object GetCheckListF2(int inspeccionId)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var inspeccion = db.Inspeccion.Find(inspeccionId);

                var noConformidadesAnteriores = db.Cumplimiento
                                                .Where(w => w.InspeccionID == inspeccion.InspeccionFase1)
                                                .Where(w => w.EvaluacionID == 3)
                                                .Select(s => s.Caracteristica);
                var noConformidadesAnterioresID = noConformidadesAnteriores.Select(s => s.ID);
                List<Norma> normas = new List<Norma>();
                var normas1 = inspeccion.InspeccionNorma.Select(s => s.Norma);
                foreach (var n in normas1)
                {
                    normas.Add(n);
                    var asociadas = n.NormasAsociadas.Select(s => s.Norma1);
                    foreach (var a in asociadas)
                    {
                        normas.Add(a);
                    }
                }
                var data =
                    normas
                        .Where(w => w.Titulo.Any(ww => ww.Requisito.Any(www => www.Caracteristica.Any(wwww => noConformidadesAnterioresID.Contains(wwww.ID)))))
                        .Select(n => new
                        {
                            Id = n.ID,
                            Text = n.Nombre,
                            Titulos = n.Titulo.Where(ww => ww.Requisito.Any(www => www.Caracteristica.Any(wwww => noConformidadesAnterioresID.Contains(wwww.ID))))
                            .Select(s => new
                            {
                                Id = s.ID,
                                Text = s.Texto,
                                Requisitos = s.Requisito
                                                  .Where(w => w.Caracteristica.Any(ww => noConformidadesAnterioresID.Contains(ww.ID)))
                                                  .Select(r => new
                                                  {
                                                      Id = r.ID,
                                                      Text = r.Descripcion,
                                                      Caracteristicas = r.Caracteristica
                                                                            .Where(w => noConformidadesAnterioresID.Contains(w.ID))
                                                                            .Select(c => new
                                                                            {
                                                                                Id = c.ID,
                                                                                Text = c.Descripcion,
                                                                                Cumplimiento = c.Cumplimiento
                                                                                                .Where(w => w.InspeccionID == inspeccionId)
                                                                                                .Select(cu => new
                                                                                                {
                                                                                                    Id = cu.EvaluacionID,
                                                                                                    Text = cu.Evaluacion.Descripcion,
                                                                                                    Observacion = cu.Observacion,
                                                                                                    Fotografia = cu.Fotografia
                                                                                                })
                                                                                                .FirstOrDefault()
                                                                            })
                                                                            .ToList()
                                                  })
                                                  .ToList()
                            })
                                .ToList()
                        })
                        .ToList();



                return new { done = true, data = data };
            }
        }
        catch (Exception ex)
        {
            log.Error("Inspecciones/GetCheckList", ex);
            return new { done = false, code = 0, message = ex.ToString() };
        }
    }
    private static object EditSpecificDataInspeccion(string info, int inspeccionId, bool all)
    {
        try
        {
            var deserializer = new JavaScriptSerializer();
            var specific = deserializer.Deserialize<List<SpecificData>>(info);
            using (var db = new CertelEntities())
            {
                var inspeccionesList = new List<Inspeccion>();
                var inspeccion = db.Inspeccion.Find(inspeccionId);

                if (all)
                {
                    inspeccionesList = db.Inspeccion
                         .Where(w => w.ServicioID == inspeccion.ServicioID)
                         .ToList();
                }
                else
                {
                    inspeccionesList.Add(inspeccion);
                }
                foreach (var i in inspeccionesList)
                {
                    foreach (var s in specific)
                    {
                        var existRecord = db.ValoresEspecificos
                                            .Where(w => w.InspeccionID == i.ID)
                                            .Where(w => w.EspecificoID == s.name)
                                            .FirstOrDefault();
                        if (existRecord != null)
                        {
                            if (s.value == string.Empty)
                            {
                                db.ValoresEspecificos.Remove(existRecord);
                                db.SaveChanges();
                            }
                            else
                            {
                                existRecord.Valor = s.value;
                                db.SaveChanges();
                            }

                            continue;
                        }
                        if (s.value == string.Empty)
                            continue;


                        var ve = new ValoresEspecificos
                        {
                            InspeccionID = i.ID,
                            EspecificoID = s.name,
                            Valor = s.value
                        };
                        db.ValoresEspecificos.Add(ve);
                        db.SaveChanges();
                    }
                }


                return new
                {
                    done = true,
                    message = "Registros Actualizados"
                };
            }



        }
        catch (Exception ex)
        {
            log.Error("Inspecciones/EditSpecificDataInspeccion", ex);
            return new { done = false, message = ex.ToString() };
        }
    }
    private static object GetInspeccion(int id)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var inspeccion = db.Inspeccion
                                    .Find(id);
                if (inspeccion == null)
                {
                    return new { done = false, message = "Inspección no existe" };
                }
                var data = new
                {
                    ItServicio = inspeccion.Servicio.IT,
                    Ubicacion = inspeccion.Ubicacion,
                    FechaInstalacion = inspeccion.FechaInstalacion.HasValue ? inspeccion.FechaInstalacion.Value.ToString("dd-MM-yyyy") : string.Empty,
                    FechaInspeccion = inspeccion.FechaInspeccion.HasValue ? inspeccion.FechaInspeccion.Value.ToString("dd-MM-yyyy") : string.Empty,
                    Aparato = inspeccion.AparatoID,
                    TipoFuncionamiento = inspeccion.TipoFuncionamientoID,
                    Destino = inspeccion.DestinoProyectoID == null ? string.Empty : inspeccion.DestinoProyecto.Descripcion,
                    DestinoId = inspeccion.DestinoProyectoID,
                    PermisoEdificacion = inspeccion.PermisoEdificacion,
                    RecepcionMunicipal = inspeccion.RecepcionMunicipal,
                    Altura = inspeccion.AlturaPisos == null ? string.Empty : inspeccion.AlturaPisos.ToString(),
                    Ingeniero = inspeccion.Ingeniero,
                    Nombre = inspeccion.NombreProyecto,
                    Edificio = inspeccion.NombreEdificio,
                    Numero = inspeccion.Numero,
                    Fec = inspeccion.FechaEmisionCertificado.HasValue ? inspeccion.FechaEmisionCertificado.Value.ToString("dd-MM-yyyy") : string.Empty,
                    Fvc = inspeccion.FechaVencimientoCertificado.HasValue ? inspeccion.FechaVencimientoCertificado.Value.ToString("dd-MM-yyyy") : string.Empty

                };
                return new
                {
                    done = true,
                    data = data
                };
            }
        }
        catch (Exception ex)
        {
            log.Error("Inspecciones/GetInspeccion", ex);
            return new { done = false, message = ex.ToString() };
        }
    }

    private static object GetSpecificDataInspeccion(int id)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var i = db.Inspeccion
                            .Where(w => w.ID == id)
                            .FirstOrDefault();
                if (i == null)
                {
                    return new { done = false, message = "Inspección no existe" };
                }
                var data = db.Especificos
                               .Select(s => new
                               {
                                   Id = s.ID,
                                   Nombre = s.Nombre,
                                   Valor = s.ValoresEspecificos
                                                .Where(w => w.InspeccionID == i.ID)
                                                .Select(ss => ss.Valor)
                                                .FirstOrDefault()
                               }).ToList();


                return new
                {
                    done = true,
                    data = data
                };
            }
        }
        catch (Exception ex)
        {
            log.Error("Inspecciones/GetInspeccion", ex);
            return new { done = false, message = ex.ToString() };
        }
    }

    private static object GetInspecciones(string sidx, string sord, int page, int rows, string it, DateTime? desde, DateTime? hasta, string us, int calificacion)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var usuario = new Encriptacion(us, false).newText;
                var user = db.Usuario.Find(usuario);
                var roles = user.UsuarioRol.Select(s => s.Rol).ToList();
                var isRevisador = roles.Contains(2) || roles.Contains(3);
                var isAprobador = roles.Contains(3);
                var defaultDate = new DateTime(2016, 01, 01);

                var list = db.Inspeccion
                                .Where(w => isRevisador ? true : w.Ingeniero == usuario)
                                .Where(w => w.Servicio.IT.Contains(it))
                                .Where(w => calificacion == -2 ? true
                                    : calificacion == -1 ? w.Calificacion == null
                                        : w.Calificacion == calificacion)
                                .Where(w => desde == null
                                    ? true
                                    : w.FechaCreacion >= desde.Value)
                                .Where(w => hasta == null
                                    ? true
                                    : w.FechaCreacion <= hasta.Value);
                int pageSize = rows;
                int totalRecords = list.Count();
                int totalPages = (int)Math.Ceiling((float)totalRecords / (float)pageSize);
                totalPages = totalPages == 0 ? 1 : totalPages;
                var result = list
                    //.OrderBy(sidx + " " + sord)
                    .OrderBy(o => o.IT)
                    .ThenBy(o => o.Fase)

                    .Skip((page - 1) * pageSize)
                    .Take(pageSize);
                var date = DateTime.Today;
                int revisar;
                var grid = new
                {
                    page = page,
                    records = totalRecords,
                    total = totalPages,
                    rows = result
                            .AsEnumerable()
                           .Select(x => new
                           {
                               Id = x.ID,
                               ItServicio = x.Servicio.IT,
                               It = x.IT,
                               Aparato = x.AparatoID == null ? string.Empty : x.Aparato.Nombre,
                               Funcionamiento = x.TipoFuncionamientoID == null ? string.Empty : x.TipoFuncionamientoAparato.Descripcion,
                               FechaCreacion = x.FechaCreacion.ToString("dd-MM-yyyy"),
                               FechaInspeccion = x.FechaInspeccion.HasValue ? x.FechaInspeccion.Value.ToString("dd-MM-yyyy") : string.Empty,
                               FechaEntrega = x.FechaEntrega.HasValue ? x.FechaEntrega.Value.ToString("dd-MM-yyyy") : string.Empty,
                               Fase = x.Fase,
                               Norma = x.InspeccionNorma.Select(s => s.Norma.Nombre).FirstOrDefault(),
                               //HasInforme = x.Informe.Any(),
                               Estado = x.EstadoInspeccion.Descripcion,
                               Ingeniero = x.Ingeniero == null ? string.Empty : x.Usuario.Nombre + " " + x.Usuario.Apellido,
                               EstadoId = x.EstadoID,
                               Revisar = revisar = !isRevisador ? 0 : x.FechaRevision.HasValue ? 1 : 2,
                               Aprobar = !isAprobador || !x.FechaRevision.HasValue ? 0 : x.FechaAprobacion.HasValue ? 1 : 2,
                               Revisar1 = revisar,
                               Aprobado = x.FechaAprobacion.HasValue,
                               AtrasadaInspeccion = date > x.FechaInspeccion && !x.Cumplimiento.Any(),
                               AtrasadaEntrega = !x.FechaEntrega.HasValue ? false : date > x.FechaEntrega.Value && x.EstadoID == 1,
                               Destinatario = x.Destinatario,
                               Fase1 = x.Fase,
                               Califica = x.Calificacion,
                               HasNextFase = x.Inspeccion1.Any() || x.CreaFaseSiguiente == true,
                               FromCotizacion = x.Servicio.CotizacionID != null
                           })
                           .ToList()
                };
                return grid;
            }

        }
        catch (Exception ex)
        {
            log.Error("ERROR AL CARGAR GRILLA DE Inspecciones", ex);
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

    public class SpecificData
    {
        public int name { get; set; }
        public string value { get; set; }
    }

}