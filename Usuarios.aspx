<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Usuarios.aspx.cs" Inherits="Usuarios" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>Certasc</title>
    <link rel="shortcut icon" href="css/images/favicon.ico" type="image/x-icon" />
    <link rel="icon" href="css/images/favicon.ico" type="image/x-icon" />
    <link href="bower_components/font-awesome/css/font-awesome.min.css" rel="stylesheet" />
    <link href="bower_components/bootstrap/dist/css/bootstrap.min.css" rel="stylesheet" />
    <link href="bower_components/jquery-ui/jquery-ui.min.css" rel="stylesheet" />
    <link href="bower_components/jquery-ui/jquery-ui.theme.min.css" rel="stylesheet" />
    <link href="bower_components/jqGrid/ui.jqgrid.css" rel="stylesheet" />
    <link href="bower_components/alertify.js/alertify.min.css" rel="stylesheet" />
    <link href="bower_components/alertify.js/bootstrap.min.css" rel="stylesheet" />
    <link href="bower_components/select2/dist/css/select2.min.css" rel="stylesheet" />
    <link href="bower_components/jquery-te-1.4.0.css" rel="stylesheet" />
    <link href="bower_components/jquery.steps/demo/css/jquery.steps.css" rel="stylesheet" />
    <link href="bower_components/bootstrap-switch.css" rel="stylesheet" />
    <link href="bower_components/magnific-popup.css" rel="stylesheet" />
    <link href="css/certel.css" rel="stylesheet" />
    <link href="css/elementos/usuarios.css?04042017" rel="stylesheet" />
</head>
<body>
    <div id="cssmenu"></div>
    <div class="container-fluid">
        <div class="row">
            <div class="col-xs-12 col-sm-12 col-md-12 col-lg-12">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h3>USUARIOS <span id="f_remove"><i class="fa fa-remove"></i></span>
                            <span id="add"><i class="fa fa-plus"></i></span>
                        </h3>
                    </div>
                    <div class="panel-body" data-height-full="true">
                        <div class="row">
                            <form id="formFiltros">
                                <div class="form-group col-xs-12 col-sm-12 col-md-5 col-lg-5">
                                    <input type="text" class="form-control" id="f_name" placeholder="Buscar usuario" />
                                </div>
                                <div class="form-group col-xs-12 col-sm-12 col-md-2 col-lg-2">
                                    <button class="btn btn-default btn-block" type="submit">
                                        <i class="fa fa-search"></i>
                                        Buscar
                                    </button>
                                </div>
                            </form>
                        </div>

                        <div class="row">
                            <table id="grid"></table>
                            <div id="pager"></div>
                        </div>

                    </div>
                </div>
            </div>
        </div>
    </div>

    <div id="add-dialog" class="dialog">
        <div class="container-fluid">
            <div class="row">
                <div class="col-xs-12 col-lg-12 col-md-12 col-sm-12">
                    <form id="add-form">
                        <div class="col-xs-12 col-lg-12 col-md-12 col-sm-12">
                            <div class="form-group">
                                <label for="au-username">Nombre de usuario (*)</label>
                                <input id="au-username" type="text" class="form-control" placeholder="Sin espacios" required maxlength="50" />
                            </div>
                        </div>
                        <div class="col-xs-12 col-lg-6 col-md-6 col-sm-12">
                            <div class="form-group">
                                <label for="au-password">Contraseña (*)</label>
                                <input id="au-password" type="password" class="form-control" placeholder="El usuario podrá cambiarla cuando lo desee" required maxlength="50" />
                            </div>
                        </div>
                        <div class="col-xs-12 col-lg-6 col-md-6 col-sm-12">
                            <div class="form-group">
                                <label for="au-password1">Repita contraseña (*)</label>
                                <input id="au-password1" type="password" class="form-control" placeholder="Repita contraseña" required maxlength="50" />
                            </div>
                        </div>
                        <div class="col-xs-12 col-lg-6 col-md-6 col-sm-12">
                            <div class="form-group">
                                <label for="au-nombre">Nombre (*)</label>
                                <input id="au-nombre" type="text" class="form-control" placeholder="Nombre" required maxlength="50" />
                            </div>
                        </div>
                        <div class="col-xs-12 col-lg-6 col-md-6 col-sm-12">
                            <div class="form-group">
                                <label for="au-apellido">Apellido (*)</label>
                                <input id="au-apellido" type="text" class="form-control" placeholder="Apellido" required maxlength="50" />
                            </div>
                        </div>
                        <div class="col-xs-12 col-lg-6 col-md-6 col-sm-12">
                            <div class="form-group">
                                <label for="au-cargo">Cargo (*)</label>
                                <input id="au-cargo" type="text" class="form-control" required />
                            </div>
                        </div>
                        <div class="col-xs-12 col-lg-6 col-md-6 col-sm-12">
                            <div class="form-group">
                                <label for="au-email">Email (*)</label>
                                <input id="au-email" type="email" class="form-control" required />
                            </div>
                        </div>
                        <div class="col-xs-12 col-lg-6 col-md-6 col-sm-12">
                            <div class="form-group">
                                <label for="au-fono">Fono (*)</label>
                                <input id="au-fono" type="text" class="form-control" required />
                            </div>
                        </div>
                        <div class="col-xs-12 col-lg-6 col-md-6 col-sm-12">
                            <div class="form-group">
                                <label for="au-celular">Celular (*)</label>
                                <input id="au-celular" type="text" class="form-control" required />
                            </div>
                        </div>
                        <div class="col-xs-12 col-lg-6 col-md-6 col-sm-12">
                            <div class="form-group">
                                <label for="au-password-mail">Contraseña Email (*)</label>
                                <input id="au-password-mail" type="password" class="form-control" placeholder="La contraseña debe ser la misma de su correo electrónico" required maxlength="50" />
                            </div>
                        </div>
                        <div class="col-xs-12 col-lg-6 col-md-6 col-sm-12">
                            <div class="form-group">
                                <label for="au-password1-mail">Repita contraseña Email (*)</label>
                                <input id="au-password1-mail" type="password" class="form-control" placeholder="Repita contraseña del correo electrónico" required maxlength="50" />
                            </div>
                        </div>

                        <div class="col-xs-12 col-lg-12 col-md-12 col-sm-12">
                            <div class="form-group">
                                <label for="au-firma">Firma</label>
                                <textarea maxlength="1000" id="au-firma" class="form-control"></textarea>
                            </div>
                        </div>
                    </form>
                </div>

            </div>
        </div>
    </div>

    <div id="edit-dialog" class="dialog">
        <div class="container-fluid">
            <div class="row">
                <div class="col-xs-12 col-lg-12 col-md-12 col-sm-12">
                    <form id="edit-form">
                        <div class="col-xs-12 col-lg-12 col-md-12 col-sm-12">
                            <div class="form-group">
                                <label for="eu-username">Nombre de usuario (*)</label>
                                <input id="eu-username" type="text" class="form-control" readonly placeholder="Sin espacios" required maxlength="50" />
                            </div>
                        </div>
                        
                        <div class="col-xs-12 col-lg-6 col-md-6 col-sm-12">
                            <div class="form-group">
                                <label for="eu-nombre">Nombre (*)</label>
                                <input id="eu-nombre" type="text" class="form-control" placeholder="Nombre" required maxlength="50" />
                            </div>
                        </div>
                        <div class="col-xs-12 col-lg-6 col-md-6 col-sm-12">
                            <div class="form-group">
                                <label for="eu-apellido">Apellido (*)</label>
                                <input id="eu-apellido" type="text" class="form-control" placeholder="Apellido" required maxlength="50" />
                            </div>
                        </div>
                        <div class="col-xs-12 col-lg-6 col-md-6 col-sm-12">
                            <div class="form-group">
                                <label for="eu-cargo">Cargo (*)</label>
                                <input id="eu-cargo" type="text" class="form-control" required />
                            </div>
                        </div>
                        <div class="col-xs-12 col-lg-6 col-md-6 col-sm-12">
                            <div class="form-group">
                                <label for="eu-email">Email (*)</label>
                                <input id="eu-email" type="email" class="form-control" required />
                            </div>
                        </div>
                        <div class="col-xs-12 col-lg-6 col-md-6 col-sm-12">
                            <div class="form-group">
                                <label for="eu-fono">Fono (*)</label>
                                <input id="eu-fono" type="text" class="form-control" required />
                            </div>
                        </div>
                        <div class="col-xs-12 col-lg-6 col-md-6 col-sm-12">
                            <div class="form-group">
                                <label for="eu-celular">Celular (*)</label>
                                <input id="eu-celular" type="text" class="form-control" required />
                            </div>
                        </div>
                        <div class="col-xs-12 col-lg-12 col-md-12 col-sm-12">
                            <div class="form-group">
                                <label for="eu-firma">Firma</label>
                                <textarea maxlength="1000" id="eu-firma" class="form-control"></textarea>
                            </div>
                        </div>
                    </form>
                </div>

            </div>
        </div>
    </div>
    <div id="roles-dialog" class="dialog">
        <div class="panel-panel-default">
            <div class="panel-body">
                <div class="form-group">
                    <div class="input-group">
                        <select id="rol" class="form-control"></select>
                        <span id="add-rol" class="input-group-addon btn btn-primary">
                            Agregar
                        </span>
              
                    </div>
                </div>
            </div>
        </div>
        <div class="container-fluid">
            <div class="row">
                <div class="col-xs-12 col-lg-12 col-md-12 col-sm-12">
                    <ul id="roles-list" class="list-group"></ul>
                </div>
            </div>
        </div>
    </div>
</body>

<script src="bower_components/jquery/dist/jquery.min.js"></script>
<script src="bower_components/bootstrap/dist/js/bootstrap.min.js"></script>
<script src="bower_components/jquery.steps/build/jquery.steps.min.js"></script>
<script src="bower_components/jquery-ui/jquery-ui.min.js"></script>
<script src="bower_components/jqGrid/jquery.jqGrid.min.js"></script>
<script src="bower_components/jqGrid/grid.locale-es.js"></script>
<script src="bower_components/alertify.js/alertify.min.js"></script>
<script src="bower_components/bootstrap-switch.js"></script>
<script src="bower_components/magnific-popup.js"></script>
<script src="bower_components/jquery-te-1.4.0.min.js"></script>
<script src="js/menu.js?20062017"></script>
<script src="js/certel.js"></script>
<script src="js/elementos/usuarios.js?20062017"></script>
</html>
