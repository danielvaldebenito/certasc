<%@ WebHandler Language="C#" Class="Login" %>

using System;
using System.Web;
using System.Linq;
using System.Web.Script.Serialization;
using System.Web.SessionState;
public class Login : IHttpHandler, IRequiresSessionState
{

    public void ProcessRequest(HttpContext context)
    {
        var data = new object();
        var serializer = new JavaScriptSerializer();
        try
        {
            var post = HttpContext.Current;
            var user = post.Request["user"];
            var pass = post.Request["pass"];
            var encrypt = new Encriptacion(pass, true);

            using (var db = new CertelEntities())
            {
                var usuario = db.Usuario
                             .Where(w => w.NombreUsuario == user)
                             .Where(w => w.Pass == encrypt.newText)
                             .Where(w => w.Habilitado == true)
                             .FirstOrDefault();

                if (usuario == null)
                    data = new { done = false, message = "Usuario y/o Contraseña inválidos" };
                else
                {
                    var encript = new Encriptacion(user.ToLower(), true);
                    var dataUser = new 
                    {
                        Usuario = encript.newText,
                        Nombre = string.Concat(usuario.Nombre, " ", usuario.Apellido),
                        Roles = usuario.UsuarioRol.Select(s => s.Rol).ToList()
                    };
                    //context.Session["dataUser"] = dataUser;
                    
                    data = new { done = true, message = "ok", user = dataUser.Usuario, name = dataUser.Nombre, roles = dataUser.Roles };
                }
            }

        }
        catch (System.Data.Entity.Core.EntityException)
        {
            data = new { done = false, message = "Está sin conexión, inténtelo más tarde" };
        }
        catch (Exception ex)
        {
            data = new { done = false, message = ex.ToString() };
        }
        finally
        {
            var json = serializer.Serialize(data);
            context.Response.ContentType = "application/json";
            context.Response.Write(json);
            context.Response.Flush();
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