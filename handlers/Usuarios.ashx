<%@ WebHandler Language="C#" Class="Usuarios" %>

using System;
using System.Web;
using System.Linq;
using System.Web.Script.Serialization;
using System.Linq.Dynamic;

public class Usuarios : IHttpHandler {

    private static readonly log4net.ILog log = log4net.LogManager.GetLogger
   (System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
    public void ProcessRequest(HttpContext context) {
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
                data = GetUsers(sidx, sord, page, rows, name);
                break;
            case "add":
                data = Add(post);
                break;
            case "edit":
                data = Edit(post);
                break;
            case "addRol":
                data = AddRol(post);
                break;
            case "removeRol":
                data = RemoveRol(post);
                break;
            case "getUser":
                data = GetUser(post);
                break;
            case "enableDisable":
                data = EnableDisable(post);
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

    private static object GetUsers(string sidx, string sord, int page, int rows, string name)
    {
        try {
            using (var db = new CertelEntities())
            {
                var list = db.Usuario
                                .Where(w => w.Nombre.Contains(name) || w.Apellido.Contains(name) || string.Concat(w.Nombre, " ", w.Apellido).Contains(name));

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
                               Username = x.NombreUsuario,
                               Nombre = x.Nombre,
                               Apellido = x.Apellido,
                               Cargo = x.Cargo,
                               Habilitado = x.Habilitado,
                               Email = x.Email,
                               Fono = x.Fono,
                               Celular = x.Celular,
                               Firma = x.Firma
                           })
                           .ToList()
                };
                return grid;
            }
        }
        catch (Exception ex) {
            return new { done = false, message = ex.Message, ex = ex.ToString() };
        }
    }
    private static object Add(HttpContext ctx)
    {
        try {
            //username, name, surname, pass, cargo
            using (var db = new CertelEntities())
            {
                var username = ctx.Request["username"];
                if (username.Contains(" "))
                    return new { done = false, message = "El nombre de usuario no puede contener espacios" };
                var exist = db.Usuario.Find(username);
                if (exist != null) {
                    return new { done = false, message = "Nombre de usuario ya existe" };
                }
                var nombre = ctx.Request["name"];
                var apellido = ctx.Request["surname"];
                var pass = ctx.Request["pass"];
                var cargo = ctx.Request["cargo"];
                var email = ctx.Request["email"];
                var firma = ctx.Request["firma"];
                var passMail = ctx.Request["passMail"];
                var fono = ctx.Request["fono"];
                var celular = ctx.Request["celular"];
                var encript = new Encriptacion(pass, true);
                var password = encript.newText;
                var encriptMail = new Encriptacion(pass, true);
                var passwordMail = encriptMail.newText;
                var user = new Usuario
                {
                    Nombre = nombre,
                    Apellido = apellido,
                    Pass = password,
                    Cargo = cargo,
                    Habilitado = true,
                    NombreUsuario = username,
                    Email = email,
                    Firma = firma,
                    PassMail = passwordMail,
                    Fono = fono,
                    Celular = celular
                };
                db.Usuario.Add(user);
                db.SaveChanges();

                return new { done = true, message = "Usuario registrado correctamente" };
            }
        }
        catch (Exception ex) {
            return new { done = false, message = ex.Message, ex = ex.ToString() };
        }
    }
    private static object Edit(HttpContext ctx)
    {
        try {
            using (var db = new CertelEntities())
            {
                var username = ctx.Request["username"];
                var exist = db.Usuario.Find(username);
                if (exist == null) {
                    return new { done = false, message = "Error: No existe usuario" };
                }
                var nombre = ctx.Request["name"];
                var apellido = ctx.Request["surname"];
                var cargo = ctx.Request["cargo"];
                var email = ctx.Request["email"];
                var firma = ctx.Request["firma"];
                var fono = ctx.Request["fono"];
                var celular = ctx.Request["celular"];
                exist.Nombre = nombre;
                exist.Apellido = apellido;
                exist.Cargo = cargo;
                exist.Email = email;
                exist.Firma = firma;
                exist.Fono = fono;
                exist.Celular = celular;
                db.SaveChanges();

                return new { done = true, message = "Usuario modificado correctamente" };
            }
        }
        catch (Exception ex) {
            return new { done = false, message = ex.Message, ex = ex.ToString() };
        }
    }
    private static object AddRol(HttpContext ctx)
    {
        try {
            using (var db = new CertelEntities())
            {
                var username = ctx.Request["username"];
                var rol = int.Parse(ctx.Request["rol"]);
                var user = db.Usuario.Find(username);
                var exists = user.UsuarioRol
                                .Where(w => w.Rol == rol)
                                .Any();
                if (exists)
                {
                    return new { done = false, message = "El usuario ya tiene el rol que intenta asignar" };
                }

                var ur = new UsuarioRol
                {
                    Usuario = username,
                    Rol = rol
                };
                db.UsuarioRol.Add(ur);
                if(rol == 1) {
                    var existPda = db.Pda.Where(w => w.Usuario == username).FirstOrDefault();
                    if(existPda == null) {
                        existPda = new Pda
                        {
                            Usuario = username
                        };
                        db.Pda.Add(existPda);
                    }
                    var uniqueNumber = user.NumeroUnico;
                    if(uniqueNumber == null) {
                        var maxUniqueNumber = db.Usuario.Max(m => m.NumeroUnico);
                        user.NumeroUnico = maxUniqueNumber + 1;

                    }
                }

                db.SaveChanges();
                return new { done = true, message = "Rol asignado correctamente" };
            }
        }
        catch (Exception ex) {
            return new { done = false, message = ex.Message, ex = ex.ToString() };
        }
    }
    private static object RemoveRol(HttpContext ctx)
    {
        try {
            using (var db = new CertelEntities())
            {
                var username = ctx.Request["username"];
                var rol = int.Parse(ctx.Request["rol"]);
                var exists = db.UsuarioRol
                                .Where(w => w.Usuario == username)
                                .Where(w => w.Rol == rol)
                                .FirstOrDefault();
                if (exists == null)
                {
                    return new { done = false, message = "El usuario no tiene el rol que intenta desasignar" };
                }

                db.UsuarioRol.Remove(exists);
                db.SaveChanges();
                return new { done = true, message = "El rol ha sido quitado del usuario" };
            }
        }
        catch (Exception ex) {
            return new { done = false, message = ex.Message, ex = ex.ToString() };
        }
    }
    private static object GetUser(HttpContext ctx)
    {
        try {
            using (var db = new CertelEntities())
            {
                var username = ctx.Request["username"];
                var user = db.Usuario.Find(username);
                if (user == null)
                    return new { done = false, message = "Usuario no existe" };
                return new
                {
                    done = true,
                    user = new
                    {
                        Username = user.NombreUsuario,
                        Name = user.Nombre,
                        Surname = user.Apellido,
                        Cargo = user.Cargo,
                        Roles = user.UsuarioRol
                                    .Select(ur => new {
                                        RolId = ur.Rol,
                                        Rol = ur.Rol1.Nombre
                                    }).ToList()
                    }
                };
            }
        }
        catch (Exception ex) {
            return new { done = false, message = ex.Message, ex = ex.ToString() };
        }
    }
    private static object EnableDisable (HttpContext ctx)
    {
        try {
            using (var db = new CertelEntities())
            {
                var username = ctx.Request["username"];
                var user = db.Usuario.Find(username);
                if (user == null)
                    return new { done = false, message = "Usuario no existe" };
                user.Habilitado = !user.Habilitado;
                db.SaveChanges();
                return new { done = true, message = "Usuario modificado correctamente" };
            }
        }
        catch (Exception ex) {
            return new { done = false, message = ex.Message, ex = ex.ToString() };
        }
    }
    public bool IsReusable {
        get {
            return true;
        }
    }

}