﻿<%@ WebHandler Language="C#" Class="ChangePassMail" %>

using System;
using System.Web;
using System.Linq;
using System.Web.Script.Serialization;
using System.Web.SessionState;
public class ChangePassMail : IHttpHandler, IRequiresSessionState {

    public void ProcessRequest (HttpContext context) {
        var post = HttpContext.Current;
        var data = new object();
        var serializer = new JavaScriptSerializer();
        var newPass = post.Request["newPass"];
        var user1 = post.Request["user"];
        var usuario = new Encriptacion(user1, false).newText;
        using (var db = new CertelEntities())
        {
            var user = db.Usuario
                         .Where(w => w.NombreUsuario == usuario)
                         .FirstOrDefault();

            if (user == null)
                data = new { code = 0, message = "No existe usuario" };
            else
            {
                var encriptNew = new Encriptacion(newPass, true);
                user.PassMail = encriptNew.newText;
                db.SaveChanges();
                data = new { code = 1, message = "Contraseña cambiada exitosamente" };
                
            }


        }
        var json = serializer.Serialize(data);
        context.Response.ContentType = "application/json";
        context.Response.Write(json);
        context.Response.Flush();

    }

    public bool IsReusable {
        get {
            return true;
        }
    }

}