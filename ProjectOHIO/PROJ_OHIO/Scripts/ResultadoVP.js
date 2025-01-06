
var __filas = 5000;

loadDataResultVP();

function loadDataResultVP() {
    var totregistros = 0;
    var paginas = 0;
    var _sw = 0;

    //alert("jim");
    var _moneda = $('#v_moneda1').val();
    var _anio = $('#v_anio1').val();
    var _mes = $('#v_mes1').val();

    if (_moneda == '') {
        _moneda = 'S';
    }

    if (_anio == 0 || _anio == '') {
        var _fecha = new Date();
        _anio = _fecha.getFullYear();
        _mes = _fecha.getMonth();
    }

    $.ajax({
        url: "/ResultadoVP/ListarResultado",
        type: "GET",
        data: {
            anio: _anio,
            mes: _mes,
            moneda: _moneda
        },
        contentType: "application/json;charset=utf-8",
        dataType: "json",
        success: function (result) {
            //console.log(result);

            $('#tabla_objeto_rvp').html("");
            $.each(JSON.parse(result), function (key, item) {
                //totregistros = item.filas;
                _sw = 1; /// para verificar que tiene registros
                //alert(item.ValorPolizaBOP);
                armarLista(item);

            });

            ///////////////////////////////////////////////////
            /////////////////PAGINADO
            ///////////////////////////////////////////////////

            /*

            paginas = Math.floor(totregistros / __filas);
            if (totregistros % __filas > 0) {
                paginas = paginas + 1;
            }

            $('#paginado_rvp').html("");
            if (paginas > 1) {
                for (var i = 1; i <= paginas; i++) {

                    $('#paginado_rvp')
                        .append($('<li class="page-item" />').html('<a class="page-link" href="#" onclick="listarpagina(\'' + '' + '\',' + i + ',' + __filas + ')">' + i + '</a>'));

                }

            } else {
                if (_sw == 0) {
                    $('#tabla_objeto_rvp')
                        .append(
                            $('<tr/>')
                                .append($('<td/>').html(""))
                                .append($('<td/>').html("No hay registros"))
                                .append($('<td/>').html(""))
                        );
                }
            }

            */

        },
        error: function (errormessage) {
            alert(errormessage.responseText);
        }
    });
}


function armarLista(item) {

    

    var _opt = '';
    //console.log(" estado : " + _estado);
    //alert(parseInt(_estado));

    //if (_estado == 1) {
    //    //alert("Verdadero");

    //    //var _opt_leer = '<a href="#" class="dropdown-item" onclick="verDatosItem(\'' + _chrcodigoorigen + '\',\'' + _numreq + '\',' + _INTITEM + ')"><i class="fas fa-eye"></i>&nbsp;&nbsp;&nbsp;ACTUALIZAR</a>';
    //    var _opt_inac = '<a href="#" class="dropdown-item" onclick="InactivarItemi(\'' + _anio + '\',\'' + _mes + '\',\'' + _nombre_mes + '\')"><i class="fas fa-window-close"></i>&nbsp;&nbsp;&nbsp;CERRAR</a>';

    //    _opt = _opt + '<div class="btn-group">';
    //    _opt = _opt + '    <button type="button" class="btn btn-default"><i class="fas fa-cog"></i></button>';
    //    _opt = _opt + '    <button type="button" class="btn btn-default dropdown-toggle dropdown-icon" data-toggle="dropdown">';
    //    _opt = _opt + '        <span class="sr-only">Toggle Dropdown</span>';
    //    _opt = _opt + '    </button>';
    //    _opt = _opt + '    <div class="dropdown-menu" role="menu">';
    //    //_opt = _opt + '        ' + _opt_leer;
    //    _opt = _opt + '        ' + _opt_inac;
    //    _opt = _opt + '    </div>';
    //    _opt = _opt + '</div>';

    //}

    //if (_estado == 0) {
    //    //alert("falso");

    //    var _opt_act = '<a href="#" class="dropdown-item" onclick="ActivarItemi(\'' + _anio + '\',\'' + _mes + '\',\'' + _nombre_mes + '\')"><i class="fas fa-key"></i>&nbsp;&nbsp;&nbsp;REAPERTURA</a>';
    //    //#D32F2F
    //    _opt = _opt + '<div class="btn-group">';
    //    _opt = _opt + '    <button type="button" class="btn btn-default bg-danger"><i class="fas fa-unlock-alt"></i></button>';
    //    _opt = _opt + '    <button type="button" class="btn btn-default dropdown-toggle dropdown-icon bg-danger" data-toggle="dropdown">';
    //    _opt = _opt + '        <span class="sr-only">Toggle Dropdown</span>';
    //    _opt = _opt + '    </button>';
    //    _opt = _opt + '    <div class="dropdown-menu" role="menu">';
    //    _opt = _opt + '        ' + _opt_act;
    //    _opt = _opt + '    </div>';
    //    _opt = _opt + '</div>';

    //}

    //alert("result  :  " + item.ValorPolizaBOP);

    var _color = "white";
    var _ValorPolizaBOP = '';
    //var _ValorPolizaBOP2 = '';

    var _icono_green = '<i class="fas fa-check-circle" style="color:green;"></i>';
    //var _icono_red = '<i class="fas fa-times-circle" style="color:red;" onclick="prueba()"></i>';

    if (item.ValorPolizaBOP == 0)
        _ValorPolizaBOP = _icono_green;
    else
        _ValorPolizaBOP = '<i class="fas fa-times-circle" style="color:red;" onclick="mostrarresultvp(\'' + item.ValorPolizaBOP_1 + '\',\'' + item.ValorPolizaBOP_2 + '\',\'' + 'Valor Poliza BOP' + '\')"></i>';

    if (item.PrimaPagada == 0)
        _PrimaPagada = _icono_green;
    else
        _PrimaPagada = '<i class="fas fa-times-circle" style="color:red;" onclick="mostrarresultvp(\'' + item.PrimaPagada_1 + '\',\'' + item.PrimaPagada_2 + '\',\'' + 'Prima Pagada sin/igv' + '\')"></i>';

    if (item.Interes == 0)
        _Interes = _icono_green;
    else
        _Interes = '<i class="fas fa-times-circle" style="color:red;" onclick="mostrarresultvp(\'' + item.Interes_1 + '\',\'' + item.Interes_2 + '\',\'' + 'Interes' + '\')"></i>';

    if (item.RescatesParciales == 0)
        _RescatesParciales = _icono_green;
    else
        _RescatesParciales = '<i class="fas fa-times-circle" style="color:red;" onclick="mostrarresultvp(\'' + item.RescatesParciales_1 + '\',\'' + item.RescatesParciales_2 + '\',\'' + 'Rescates Parciales' + '\')"></i>';

    if (item.CargoPorRescateParcial == 0)
        _CargoPorRescateParcial = _icono_green;
    else
        _CargoPorRescateParcial = '<i class="fas fa-times-circle" style="color:red;" onclick="mostrarresultvp(\'' + item.CargoPorRescateParcial_1 + '\',\'' + item.CargoPorRescateParcial_2 + '\',\'' + 'Cargo por Rescate Parcial' + '\')"></i>';

    if (item.GastoAdministrativoPB == 0)
        _GastoAdministrativoPB = _icono_green;
    else
        _GastoAdministrativoPB = '<i class="fas fa-times-circle" style="color:red;" onclick="mostrarresultvp(\'' + item.GastoAdministrativoPB_1 + '\',\'' + item.GastoAdministrativoPB_2 + '\',\'' + 'Gasto administrativo PB' + '\')"></i>';

    if (item.GastoAdministrativoPexc == 0)
        _GastoAdministrativoPexc = _icono_green;
    else
        _GastoAdministrativoPexc = '<i class="fas fa-times-circle" style="color:red;" onclick="mostrarresultvp(\'' + item.GastoAdministrativoPexc_1 + '\',\'' + item.GastoAdministrativoPexc_2 + '\',\'' + 'Gasto administrativo Pexc' + '\')"></i>';

    if (item.GastoAdministrativoDisminucionSA == 0)
        _GastoAdministrativoDisminucionSA = _icono_green;
    else
        _GastoAdministrativoDisminucionSA = '<i class="fas fa-times-circle" style="color:red;" onclick="mostrarresultvp(\'' + item.GastoAdministrativoDisminucionSA_1 + '\',\'' + item.GastoAdministrativoDisminucionSA_2 + '\',\'' + 'G.Adm Disminucion SA' + '\')"></i>';

    if (item.COI_Muerte == 0)
        _COI_Muerte = _icono_green;
    else
        _COI_Muerte = '<i class="fas fa-times-circle" style="color:red;" onclick="mostrarresultvp(\'' + item.COI_Muerte_1 + '\',\'' + item.COI_Muerte_2 + '\',\'' + 'COI MUERTE' + '\')"></i>';

    if (item.COI_ITP == 0)
        _COI_ITP = _icono_green;
    else
        _COI_ITP = '<i class="fas fa-times-circle" style="color:red;" onclick="mostrarresultvp(\'' + item.COI_ITP_1 + '\',\'' + item.COI_ITP_2 + '\',\'' + 'COI  ITP' + '\')"></i>';

    if (item.COI_MuerteAccidental == 0)
        _COI_MuerteAccidental = _icono_green;
    else
        _COI_MuerteAccidental = '<i class="fas fa-times-circle" style="color:red;" onclick="mostrarresultvp(\'' + item.COI_MuerteAccidental_1 + '\',\'' + item.COI_MuerteAccidental_2 + '\',\'' + 'COI  Muerte Accidental' + '\')"></i>';

    if (item.COI_EnfermedadesGraves == 0)
        _COI_EnfermedadesGraves = _icono_green;
    else
        _COI_EnfermedadesGraves = '<i class="fas fa-times-circle" style="color:red;" onclick="mostrarresultvp(\'' + item.COI_EnfermedadesGraves_1 + '\',\'' + item.COI_EnfermedadesGraves_2 + '\',\'' + 'COI Enfermedades Graves' + '\')"></i>';

    if (item.COI_Exoneracion == 0)
        _COI_Exoneracion = _icono_green;
    else
        _COI_Exoneracion = '<i class="fas fa-times-circle" style="color:red;" onclick="mostrarresultvp(\'' + item.COI_Exoneracion_1 + '\',\'' + item.COI_Exoneracion_2 + '\',\'' + 'COI Exoneracion' + '\')"></i>';

    if (item.COI_MuerteAA == 0)
        _COI_MuerteAA = _icono_green;
    else
        _COI_MuerteAA = '<i class="fas fa-times-circle" style="color:red;" onclick="mostrarresultvp(\'' + item.COI_MuerteAA_1 + '\',\'' + item.COI_MuerteAA_2 + '\',\'' + 'COI Muerte AA' + '\')"></i>';

    if (item.COI_ITP_AA == 0)
        _COI_ITP_AA = _icono_green;
    else
        _COI_ITP_AA = '<i class="fas fa-times-circle" style="color:red;" onclick="mostrarresultvp(\'' + item.COI_ITP_AA_1 + '\',\'' + item.COI_ITP_AA_2 + '\',\'' + 'COI ITP AA' + '\')"></i>';

    if (item.CargoPorRescateTotalCancelacion == 0)
        _CargoPorRescateTotalCancelacion = _icono_green;
    else
        _CargoPorRescateTotalCancelacion = '<i class="fas fa-times-circle" style="color:red;" onclick="mostrarresultvp(\'' + item.CargoPorRescateTotalCancelacion_1 + '\',\'' + item.CargoPorRescateTotalCancelacion_2 + '\',\'' + 'Cargo por Rescate Total - Cancelación' + '\')"></i>';

    if (item.COIS_PendientesPago == 0)
        _COIS_PendientesPago = _icono_green;
    else
        _COIS_PendientesPago = '<i class="fas fa-times-circle" style="color:red;" onclick="mostrarresultvp(\'' + item.COIS_PendientesPago_1 + '\',\'' + item.COIS_PendientesPago_2 + '\',\'' + 'COIS Pendientes de Pago' + '\')"></i>';

    if (item.ValorPolizaEOP == 0)
        _ValorPolizaEOP = _icono_green;
    else
        _ValorPolizaEOP = '<i class="fas fa-times-circle" style="color:red;" onclick="mostrarresultvp(\'' + item.ValorPolizaEOP_1 + '\',\'' + item.ValorPolizaEOP_2 + '\',\'' + 'Valor Poliza EOP' + '\')"></i>';

    if (item.ValorRescate == 0)
        _ValorRescate = _icono_green;
    else
        _ValorRescate = '<i class="fas fa-times-circle" style="color:red;" onclick="mostrarresultvp(\'' + item.ValorRescate_1 + '\',\'' + item.ValorRescate_2 + '\',\'' + 'Valor de Rescate' + '\')"></i>';


    $('#tabla_objeto_rvp')
        .append(
            $('<tr style="background-color:' + _color + ';"/>')
                //.append($('<td/>').html(_opt))
                //// item.NumeroPoliza
                .append($('<td/>').html(item.NumeroPoliza))
                .append($('<td/>').html(_ValorPolizaBOP))
                .append($('<td/>').html(_PrimaPagada))
                .append($('<td/>').html(_Interes))
                .append($('<td/>').html(_RescatesParciales))
                .append($('<td/>').html(_CargoPorRescateParcial))
                .append($('<td/>').html(_GastoAdministrativoPB))
                .append($('<td/>').html(_GastoAdministrativoPexc))
                .append($('<td/>').html(_GastoAdministrativoDisminucionSA))
                .append($('<td/>').html(_COI_Muerte))
                .append($('<td/>').html(_COI_ITP))
                .append($('<td/>').html(_COI_MuerteAccidental))
                .append($('<td/>').html(_COI_EnfermedadesGraves))
                .append($('<td/>').html(_COI_Exoneracion))
                .append($('<td/>').html(_COI_MuerteAA))
                .append($('<td/>').html(_COI_ITP_AA))
                .append($('<td/>').html(_CargoPorRescateTotalCancelacion))
                .append($('<td/>').html(_COIS_PendientesPago))
                .append($('<td/>').html(_ValorPolizaEOP))
                .append($('<td/>').html(_ValorRescate))
        );

}


function mostrarresultvp(valor_excel, resultado, texto) {

    $('#titulo_data').html(texto);
    $('#v_excel').val(valor_excel);
    $('#v_calculo').val(resultado);

    $('#modal-sm-result-vp').modal('show');
}


function export_excel_g() {

    var _moneda = $("#v_moneda1").val();
    var _anio = $("#v_anio1").val();
    var _mes = $("#v_mes1").val();
    
    //if (isNaN(_documento)) {
    //    _documento = 0;
    //}

    //if (isNaN(_func_emite)) {
    //    _func_emite = 0;
    //}
    //if (isNaN(_func_destino)) {
    //    _func_destino = 0;
    //}

    $.ajax({

        cache: false,
        url: '/ResultadoVP/Exportarconsultageneral',
        data: {
            anio: _anio,
            mes: _mes,
            moneda: _moneda
        },
        //contentType: "application/json;charset=utf-8",
        dataType: "json",
        success: function (data) {
            //alert("ingreso");
            //console.log(data.FileGuid);
            /*var response = JSON.stringify(data);*/
            //console.log('**********');
            //console.log(response);

            window.location = '/ResultadoVP/Download?fileGuid=' + data.FileGuid
                + '&filename=' + data.FileName;
        },
        error: function (errormessage) {
            //toastr.error(errormessage.responseText);
            console.log("consulta 5");
        }
    })

}

