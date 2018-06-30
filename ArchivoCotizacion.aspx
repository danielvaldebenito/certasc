<%@ Page Language="C#" AutoEventWireup="true" CodeFile="ArchivoCotizacion.aspx.cs" Inherits="ArchivoCotizacion" %>

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
    <link href="css/certel.css" rel="stylesheet" />
    <link href="css/elementos/archivosCotizacion.css" rel="stylesheet" />
</head>
<body>
     <div class="container-fluid">
        <div class="row">
            <div class="col-xs-12">
                <div class="page-header">
                    <h1 id="titulo">ARCHIVOS DE LA COTIZACION</h1>
                </div>
            </div>
        </div>

        <div class="row">
            <div class="col-xs-12">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h3>CERTEL</h3>
                    </div>

                    <div class="panel-body">
                      
                        <div class="row" id="filesPanel">
                              <h3 id="tituloPanel"></h3>
                        </div>
                    </div>

                </div>
            </div>
        </div>



    </div>



</body>
<script src="bower_components/jquery/dist/jquery.min.js"></script>
<script src="bower_components/bootstrap/dist/js/bootstrap.min.js"></script>
<script src="bower_components/jquery-ui/jquery-ui.min.js"></script>
<script src="js/elementos/archivosCotizacion.js?1=15032018"></script>
</html>
