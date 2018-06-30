<%@ Page Language="C#" AutoEventWireup="true" CodeFile="calendar.aspx.cs" Inherits="calendar" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>Certel</title>
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
    <link href="js/full-calendar/fullcalendar.min.css" rel="stylesheet" />
    <link href="js/full-calendar/fullcalendar.print.min.css" rel="stylesheet" media='print' />
    <link href="css/elementos/calendar.css" rel="stylesheet" />
</head>
<body>
    <div id="cssmenu"></div>

    <div class="container-fluid">
        <div class="row">
            <div class="col-xs-12">
                <div class="page-header">
                    <h1 id="titulo">Calendario de Actividades</h1>
                </div>
            </div>



            <div class="col-xs-12">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h3>Ingeniero</h3>
                    </div>
                    <div class="panel-body">
                        <select class="form-control" id="cbxIngeniero">
                            <option value="0">Seleccione...</option>
                        </select>
                    </div>
                </div>

            </div>



            <div class="col-xs-12">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h3>Actividades</h3>
                    </div>
                    <div class="panel-body" id="calendar-container">
                        <div id='calendar'></div>
                    </div>
                </div>
                
            </div>

        </div>

    </div>
    <%--Lock Screen--%>
    <div class="lock-screen">
        <i class="fa fa-cog fa-spin fa-5x fa-fw margin-bottom"></i>
        <span class="sr-only">Cargando...</span>
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
<script src="js/moment.js"></script>
<script src="js/full-calendar/fullcalendar.min.js"></script>
<script src="js/elementos/calendar.js?1=070320182335"></script>
<script src="js/full-calendar/locale-all.js"></script>
</html>
