<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Clientes.aspx.cs" Inherits="Clientes" %>

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
    <link href="bower_components/jqGrid/ui.jqgrid.css" rel="stylesheet" />
    <link href="bower_components/alertify.js/alertify.min.css" rel="stylesheet" />
    <link href="bower_components/alertify.js/bootstrap.min.css" rel="stylesheet" />
    <link href="bower_components/select2/dist/css/select2.min.css" rel="stylesheet" />
    <link href="bower_components/jquery.steps/demo/css/jquery.steps.css" rel="stylesheet" />
    <link href="bower_components/bootstrap-switch.css" rel="stylesheet" />
    <link href="bower_components/magnific-popup.css" rel="stylesheet" />
    <link href="css/certel.css" rel="stylesheet" />
    <link href="css/elementos/clientes.css?04042017" rel="stylesheet" />
</head>
<body>
    <div id="cssmenu"></div>
    <div class="container-fluid">
        <div class="row">
            <div class="col-xs-12 col-sm-12 col-md-12 col-lg-12">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h3>CLIENTES <span id="f_remove"><i class="fa fa-remove"></i></span>
                            <span id="add-client"><i class="fa fa-plus"></i></span>
                        </h3>
                    </div>
                    <div class="panel-body" data-height-full="true">
                        <div class="row">
                            <form id="formFiltros">
                                <div class="form-group col-xs-12 col-sm-12 col-md-5 col-lg-5">
                                    <input type="text" class="form-control" id="f_rut" placeholder="Rut Cliente" />
                                </div>
                                <div class="form-group col-xs-12 col-sm-12 col-md-5 col-lg-5">
                                    <input type="text" class="form-control" id="f_name" placeholder="Nombre Cliente" />
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

    <%--Add--%>

    <div id="add-client-dialog" class="dialog">
        <div class="container-fluid">
            <div class="row">
                <div class="col-xs-12 col-md-12 col-sm-12 col lg-12">
                    <form id="add-client-form">
                        <div class="form-group">
                            <label for="ac_rut">Rut (*)</label>
                            <input type="text" value="" class="form-control" id="ac_rut" required maxlength="50" />
                        </div>
                        <div class="form-group">
                            <label for="ac_nombre">Nombre (*)</label>
                            <input type="text" value="" class="form-control" id="ac_nombre" required maxlength="50"/>
                        </div>
                        <div class="form-group">
                            <label for="ac_giro">Giro</label>
                            <input type="text" value="" class="form-control" id="ac_giro" maxlength="50"/>
                        </div>
                        <div class="form-group">
                            <label for="ac_direccion">Dirección</label>
                            <input type="text" value="" class="form-control" id="ac_direccion" maxlength="200" />
                        </div>
                        <div class="form-group">
                            <label for="ac_telefono">Teléfono Contacto</label>
                            <input type="text" value="" class="form-control" id="ac_telefono" />
                        </div>
                        <div class="form-group">
                            <label for="ac_mail">Email Contacto</label>
                            <input type="email" value="" class="form-control" id="ac_mail" />
                        </div>
                    </form>
                </div>
            </div>
        </div>
    </div>

    <%--Edit--%>

    <div id="edit-client-dialog" class="dialog">
        <div class="container-fluid">
            <div class="row">
                <div class="col-xs-12 col-md-12 col-sm-12 col lg-12">
                    <form id="edit-client-form">
                        <div class="form-group">
                            <label for="ac_rut">Rut (*)</label>
                            <input type="text" value="" class="form-control" id="ec_rut" required maxlength="50" />
                        </div>
                        <div class="form-group">
                            <label for="ac_nombre">Nombre (*)</label>
                            <input type="text" value="" class="form-control" id="ec_nombre" required maxlength="50"/>
                        </div>
                        <div class="form-group">
                            <label for="ac_giro">Giro</label>
                            <input type="text" value="" class="form-control" id="ec_giro" maxlength="50"/>
                        </div>
                        <div class="form-group">
                            <label for="ac_direccion">Dirección</label>
                            <input type="text" value="" class="form-control" id="ec_direccion" maxlength="200" />
                        </div>
                        <div class="form-group">
                            <label for="ac_telefono">Teléfono Contacto</label>
                            <input type="text" value="" class="form-control" id="ec_telefono" />
                        </div>
                        <div class="form-group">
                            <label for="ac_mail">Email Contacto</label>
                            <input type="email" value="" class="form-control" id="ec_mail" />
                        </div>
                    </form>
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
<script src="js/menu.js"></script>
<script src="js/certel.js"></script>
<script src="js/jquery.Rut.js"></script>
<script src="js/elementos/clientes.js?16032017"></script>
</html>
