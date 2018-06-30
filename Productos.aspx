<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Productos.aspx.cs" Inherits="Productos" %>

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
    <link href="bower_components/jquery.steps/demo/css/jquery.steps.css" rel="stylesheet" />
    <link href="bower_components/bootstrap-switch.css" rel="stylesheet" />
    <link href="bower_components/magnific-popup.css" rel="stylesheet" />
    <link href="css/certel.css" rel="stylesheet" />
</head>
<body>
    <div id="cssmenu"></div>
    <div class="col-xs-12 col-sm-12 col-md-3 col-lg-3">
        <div class="panel panel-default">
            <div class="panel-heading">
                <h3>filtros <span id="f_remove"><i class="fa fa-times"></i></span></h3>
            </div>
            <div class="panel-body">
                <form id="filter-form">
                    <div class="form-group">
                        <label>Nombre</label>
                        <input id="f_name" type="text" placeholder="Buscar por nombre" class="form-control" />
                    </div>
                    <div class="form-group">
                        <button type="submit" class="btn btn-default btn-block">
                            <i class="fa fa-search"></i> Buscar
                        </button>
                    </div>
                </form>
            </div>
        </div>
    </div>
    <div class="col-xs-12 col-sm-12 col-md-9 col-lg-9">
        <div class="panel panel-default">
            <div class="panel-heading">
                <h3>Productos
                 <span id="add"><i class="fa fa-plus"></i></span>
                </h3>
            </div>
            <div class="panel-body">
                <div class="row">
                    <table id="grid"></table>
                    <div id="pager"></div>
                </div>

            </div>
        </div>
    </div>



    <div id="add-dialog" class="dialog">
        <div class="container-fluid">
            <div class="row">
                <div class="col-xs-12">
                    <form id="add-form">

                        <div class="form-group">
                            <label for="add-name">Nombre (*)</label>
                            <input id="add-name" type="text" placeholder="Nombre" value="" class="form-control" />
                        </div>
                        
                    </form>
                </div>
            </div>
        </div>
    </div>


    <div id="edit-dialog" class="dialog">
        <div class="container-fluid">
            <div class="row">
                <div class="col-xs-12">
                    <form id="edit-form">

                        <div class="form-group">
                            <label for="edit-name">Nombre (*)</label>
                            <input id="edit-name" type="text" placeholder="Nombre" value="" class="form-control" />
                        </div>
                        
                    </form>
                </div>
            </div>
        </div>
    </div>
</body>
<script src="bower_components/jquery/dist/jquery.min.js"></script>
<script src="bower_components/bootstrap/dist/js/bootstrap.min.js"></script>
<script src="bower_components/jquery-ui/jquery-ui.min.js"></script>
<script src="bower_components/jqGrid/jquery.jqGrid.min.js"></script>
<script src="bower_components/jqGrid/grid.locale-es.js"></script>
<script src="bower_components/alertify.js/alertify.min.js"></script>
<script src="bower_components/bootstrap-switch.js"></script>
<script src="bower_components/magnific-popup.js"></script>
<script src="js/menu.js"></script>
<script src="js/certel.js"></script>
<script src="js/elementos/productos.js"></script>
</html>
