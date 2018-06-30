<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Login.aspx.cs" Inherits="Login" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>Certasc</title>
    <link rel="shortcut icon" href="css/images/favicon.ico" type="image/x-icon" />
    <link rel="icon" href="css/images/favicon.ico" type="image/x-icon" />
    <link href="bower_components/font-awesome/css/font-awesome.min.css" rel="stylesheet" />
    <link href="bower_components/bootstrap/dist/css/bootstrap.min.css" rel="stylesheet" />
    <link href="bower_components/jquery-ui/jquery-ui.min.css" rel="stylesheet"/>
    <link href="bower_components/jquery-ui/jquery-ui.theme.min.css" rel="stylesheet" />
    <link href="bower_components/alertify.js/alertify.min.css" rel="stylesheet" />
    <link href="bower_components/alertify.js/bootstrap.min.css" rel="stylesheet" />
    <link href="css/certel.css" rel="stylesheet" />
    <link href="css/elementos/login.css" rel="stylesheet"/>
</head>
<body>
    <div class="container-fluid">
        <div class="row">
            <div class="col-lg-4 col-xs-12 col-sm-3 col-md-3"></div>
            <div class="col-lg-4 col-xs-12 col-sm-6 col-md-6 text-center">
                <img class="img-responsive img-thumbnail" src="css/images/logo.png" alt="Certel" width="300" />
            </div>
            <div class="col-lg-4 col-xs-12 col-sm-3 col-md-3"></div>
        </div>
        <div class="row">
            <div class="col-lg-4 col-xs-12 col-sm-3 col-md-3"></div>
            <div class="col-lg-4 col-xs-12 col-sm-6 col-md-6">
                <div class="panel-default">
                    <div class="panel-heading">
                        <h3>Bienvenido a Certasc</h3>
                    </div>
                    <div class="panel-body">
                        <form id="form-login">
                            <div class="form-group">
                                <label for="usuario">Usuario</label>
                                <div class="input-group">
                                    <span class="input-group-addon">
                                        <i class="fa fa-user"></i>
                                    </span>
                                    <input type="text" id="user" placeholder="Usuario" maxlength="50" class="form-control" />
                                </div>
                            </div>
                            <div class="form-group">
                                <label for="password">Contraseña</label>
                                <div class="input-group">
                                    <span class="input-group-addon">
                                        <i class="fa fa-unlock-alt"></i>
                                    </span>
                                    <input type="password" id="pass" placeholder="Contraseña" maxlength="50" class="form-control" />
                                </div>
                            </div>
                            <button type="submit" class="btn btn-primary btn-block btn-lg"><i class="fa fa-sign-in"></i> Entrar</button>
                        </form>
                    </div>
                </div>
                
            </div>
            <div class="col-lg-4 col-xs-12 col-sm-3 col-md-3"></div>
        </div>
    </div>
     <%--Foot-Page--%>
    <div class="foot-page">Certel - Certificación en Elevación S.A.</div>

    <%--Lock Screen--%>
    <div class="lock-screen dialog">
        <i class="fa fa-cog fa-spin fa-3x fa-fw margin-bottom"></i>
        <span class="sr-only">Cargando...</span>
    </div>
</body>
<script src="bower_components/jquery/dist/jquery.min.js"></script>
<script src="bower_components/bootstrap/dist/js/bootstrap.min.js"></script>
<script src="bower_components/jquery-ui/jquery-ui.min.js"></script>
<script src="bower_components/jqGrid/jquery.jqGrid.min.js"></script>
<script src="bower_components/jqGrid/grid.locale-es.js"></script>
<script src="bower_components/alertify.js/alertify.min.js"></script>
<script src="js/certel.js"></script>
<script src="js/elementos/login.js"></script>
</html>
