
/*
function CalcularVTMoce(_anio, _mes, _nombremes) {
    $('#modal-sm-procesar_vt_moce').modal('show');
}
*/


var __filas = 15;

loadDataTemporal();

function loadDataTemporal() {
    var totregistros = 0;
    var paginas = 0;

    var _busqueda = "";
    var _sw = 0;

    $.ajax({
        url: "/Temporal/List",
        type: "GET",
        data: {
            busqueda: _busqueda,
            estado: 1,
            page: 1,
            filas: __filas
        },
        contentType: "application/json;charset=utf-8",
        dataType: "json",
        success: function (result) {
            //console.log(result);

            $('#tabla_objeto_tem').html("");
            $.each(JSON.parse(result), function (key, item) {
                totregistros = item.filas;
                _sw = 1; /// para verificar que tiene registros

                //alert(item.mes_vida_temporal);
                armarLista(item.item, item.anio_vida_temporal, item.mes_vida_temporal, item.nombre_mes, item.FlagDataPolizasSoles, item.FlagDataPolizasDolares, item.estado_periodo);

            });

            ///////////////////////////////////////////////////
            /////////////////PAGINADO
            ///////////////////////////////////////////////////
            paginas = Math.floor(totregistros / __filas);
            if (totregistros % __filas > 0) {
                paginas = paginas + 1;
            }

            $('#paginado_tem').html("");
            if (paginas > 1) {
                for (var i = 1; i <= paginas; i++) {

                    $('#paginado_tem')
                        .append($('<li class="page-item" />').html('<a class="page-link" href="#" onclick="listarpagina(\'' + _busqueda + '\',' + i + ',' + __filas + ')">' + i + '</a>'));

                }

            } else {
                if (_sw == 0) {
                    $('#tabla_objeto_tem')
                        .append(
                            $('<tr/>')
                                .append($('<td/>').html(""))
                                .append($('<td/>').html("No hay registros"))
                                .append($('<td/>').html(""))
                        );
                }
            }

        },
        error: function (errormessage) {
            alert(errormessage.responseText);
        }
    });
}


function listarpagina(_busqueda, _pagina, _filas) {

    $.ajax({
        url: "/Temporal/List",
        type: "GET",
        data: {
            busqueda: _busqueda,
            estado: 1,
            page: _pagina,
            filas: _filas
        },
        contentType: "application/json;charset=utf-8",
        dataType: "json",
        success: function (result) {
            console.log(result);

            $('#tabla_objeto_tem').html("");
            $.each(JSON.parse(result), function (key, item) {

                armarLista(item.item, item.anio_vida_temporal, item.mes_vida_temporal, item.nombre_mes, item.FlagDataPolizasSoles, item.FlagDataPolizasDolares, item.estado_periodo);

            });

        },
        error: function (errormessage) {
            alert(errormessage.responseText);
        }
    });

}

function armarLista(_item, _anio, _mes, _nombre_mes, _FlagDataPolizasSoles, _FlagDataPolizasDolares, _activo) {

    var var__FlagDataPolizasSoles = '';
    var var__FlagDataPolizasDolares = '';
    //var var__FlagDataCobranzasSoles = '';
    //var var__FlagDataCobranzasDolares = '';
    //var var__FlagDataEndoso = '';


    if (_FlagDataPolizasSoles == 0)
        var__FlagDataPolizasSoles = 'No';
    else
        var__FlagDataPolizasSoles = 'Si';

    if (_FlagDataPolizasDolares == 0)
        var__FlagDataPolizasDolares = 'No';
    else
        var__FlagDataPolizasDolares = 'Si';

    //if (_FlagDataCobranzasSoles == 0)
    //    var__FlagDataCobranzasSoles = 'No';
    //else
    //    var__FlagDataCobranzasSoles = 'Si';

    //if (_FlagDataCobranzasDolares == 0)
    //    var__FlagDataCobranzasDolares = 'No';
    //else
    //    var__FlagDataCobranzasDolares = 'Si';

    //if (_FlagDataEndoso == 0)
    //    var__FlagDataEndoso = 'No';
    //else
    //    var__FlagDataEndoso = 'Si';


    var _opt = '';
    var _anio_html = '';
    var _mes_html = '';

    //var _opt_result_valor_poliza = '<a href="#" class="dropdown-item" onclick="IrResultadoVP(\'' + _anio + '\',\'' + _mes + '\')"><i class="fas fa-clone"></i>&nbsp;&nbsp;&nbsp;Resultados Valor P&oacute;liza</a>';
    var _opt_result_vida_univ = '<a href="#" class="dropdown-item" onclick="IrResultadoVT(\'' + _anio + '\',\'' + _mes + '\')"><i class="fas fa-clone"></i>&nbsp;&nbsp;&nbsp;Resultados Reserva Matem&aacute;tica VT</a>';

    if (_activo == 1) {
        //alert("ACTIVO");
        _anio_html = '<font color="#000000">' + _anio + '</font>';
        _mes_html = '<font color="#000000">' + _nombre_mes + '</font>';

        var _opt_polizas = '<a href="#" class="dropdown-item" onclick="mostrar_form_polizas(\'' + _anio + '\',\'' + _mes + '\',\'' + _nombre_mes + '\')"><i class="fas fa-database"></i>&nbsp;&nbsp;&nbsp;Cargar Data de P&oacute;lizas</a>';
        //var _opt_cobranzas = '<a href="#" class="dropdown-item" onclick="mostrar_form_cobranzas(\'' + _anio + '\',\'' + _mes + '\',\'' + _nombre_mes + '\')"><i class="fas fa-database"></i>&nbsp;&nbsp;&nbsp;Cargar Data de Cobranzas</a>';
        //var _opt_endosos = '<a href="#" class="dropdown-item" onclick="mostrar_form_endosos(\'' + _anio + '\',\'' + _mes + '\',\'' + _nombre_mes + '\')"><i class="fas fa-database"></i>&nbsp;&nbsp;&nbsp;Cargar Data de Endosos</a>';
        //var _opt_valor_poliza = '<a href="#" class="dropdown-item" onclick="CalcularValorPoliza(\'' + _anio + '\',\'' + _mes + '\',\'' + _nombre_mes + '\')"><i class="fas fa-cogs"></i>&nbsp;&nbsp;&nbsp;Procesar Valor P&oacute;liza</a>';
        var _opt_vida_univ = '<a href="#" class="dropdown-item" onclick="CalcularReserva(\'' + _anio + '\',\'' + _mes + '\',\'' + _nombre_mes + '\')"><i class="fas fa-cogs"></i>&nbsp;&nbsp;&nbsp;Procesar Reserva Matem&aacute;tica VT</a>';
        var _opt_vida_univ_moce = '<a href="#" class="dropdown-item" onclick="CalcularVTMoce(\'' + _anio + '\',\'' + _mes + '\',\'' + _nombre_mes + '\')"><i class="fas fa-cogs"></i>&nbsp;&nbsp;&nbsp;Procesar Reserva Matem&aacute;tica MOCE</a>';

        _opt = _opt + '<div class="btn-group">';
        _opt = _opt + '    <button type="button" class="btn btn-default"><i class="fas fa-cog"></i></button>';
        _opt = _opt + '    <button type="button" class="btn btn-default dropdown-toggle dropdown-icon" data-toggle="dropdown">';
        _opt = _opt + '        <span class="sr-only">Toggle Dropdown</span>';
        _opt = _opt + '    </button>';
        _opt = _opt + '    <div class="dropdown-menu" role="menu">';
        _opt = _opt + '        ' + _opt_polizas;
        //_opt = _opt + '        ' + _opt_cobranzas;
        //_opt = _opt + '        ' + _opt_endosos;
        //_opt = _opt + '        ' + _opt_valor_poliza;
        _opt = _opt + '        ' + _opt_vida_univ;
        //_opt = _opt + '        ' + _opt_result_valor_poliza;
        _opt = _opt + '        ' + _opt_result_vida_univ;
        _opt = _opt + '        ' + _opt_vida_univ_moce;
        _opt = _opt + '    </div>';
        _opt = _opt + '</div>';

    }

    if (_activo == 0) {
        // #c0392b
        _anio_html = '<font color="#ffa000">' + _anio + '</font>';
        _mes_html = '<font color="#ffa000">' + _nombre_mes + '</font>';

        //var _opt_result_valor_poliza = '<a href="#" class="dropdown-item" onclick=""><i class="fas fa-clone"></i>&nbsp;&nbsp;&nbsp;Resultados Valor P&oacute;liza</a>';
        //var _opt_result_vida_univ = '<a href="#" class="dropdown-item" onclick=""><i class="fas fa-clone"></i>&nbsp;&nbsp;&nbsp;Resultados Reserva Matematica VU</a>';

        _opt = _opt + '<div class="btn-group">';
        _opt = _opt + '    <button type="button" class="btn btn-default bg-danger"><i class="fas fa-unlock-alt"></i></button>';
        _opt = _opt + '    <button type="button" class="btn btn-default dropdown-toggle dropdown-icon bg-danger" data-toggle="dropdown">';
        _opt = _opt + '        <span class="sr-only">Toggle Dropdown</span>';
        _opt = _opt + '    </button>';
        _opt = _opt + '    <div class="dropdown-menu" role="menu">';
        //_opt = _opt + '        ' + _opt_result_valor_poliza;
        _opt = _opt + '        ' + _opt_result_vida_univ;
        _opt = _opt + '    </div>';
        _opt = _opt + '</div>';

    }

    var _color = "";



    $('#tabla_objeto_tem')
        .append(
            $('<tr style="background-color:' + _color + ';"/>')
                .append($('<td/>').html(_opt))
                .append($('<td/>').html(_anio_html))
                .append($('<td/>').html(_mes_html))
                .append($('<td/>').html(var__FlagDataPolizasSoles))
                .append($('<td/>').html(var__FlagDataPolizasDolares))                
        );


}


function CalcularVTMoce(_anio, _mes, _nombremes) {
    //$('#modal-sm-procesar_vt_moce').modal('show');

    $("#titulo_modal_moce_vt").html('Procesar Reserva Matemática V.T. <b style="color: black;">&nbsp;&nbsp;&nbsp;&nbsp;(' + _nombremes + '-' + _anio + ')</b>');

    $("#v_anio_moce_vt").val(_anio);
    $("#v_mes_moce_vt").val(_mes);

    $("#v_moneda_moce_vt").val("");
    $("#v_moneda_moce_vt").change();

    $('#modal-sm-procesar_vt_moce').modal('show');

}


function CalcularReserva(_anio, _mes, _nombremes) {

    $("#titulo_modal_vt").html('Procesar Reserva Matemática V.T. <b style="color: black;">&nbsp;&nbsp;&nbsp;&nbsp;(' + _nombremes + '-' + _anio + ')</b>');

    $("#v_anio_reserva_vt").val(_anio);
    $("#v_mes_reserva_vt").val(_mes);

    $("#v_moneda_reserva_vt").val("");
    $("#v_moneda_reserva_vt").change();

    $('#modal-sm-procesar_vt').modal('show');

}


async function limpiarFormularioTemporal() {

    //alert("001");
    loadCombo(0, 'Temporal', 'ListPeriodos', 'v_periodo_temporal', 1);
    //alert("002");
    $('#modal-sm-temporal').modal('show');
    //alert("003");

}


function mostrar_form_polizas(_anio, _mes, _mesnombre) {

    $("#titulo_modal_poliza").html('Cargar Data de P&oacute;lizas <b style="color: black;">&nbsp;&nbsp;&nbsp;&nbsp;(Periodo ' + _anio + '-' + _mesnombre + ')</b>');

    $("#v_anio_polizas").val(_anio);
    $("#v_mes_polizas").val(_mes);

    $("#v_moneda_poliza").val("");
    $("#v_moneda_poliza").change();

    $("#customFile").val("");

    $('#modal-sm-polizas').modal('show');
}


function GuardarTemporal() {
    var id = $("#v_id").val();
    var _periodo = $("#v_periodo_temporal").val();

    if ($.trim(_periodo) == '') {
        toastr.error("Ingrese cantidad");
        return false;
    }

    if (id == "0") {
        grabaitem('Guardar');
    } else {
        grabaitem('Actualizar');
    }

}


function grabaitem(_accion) {

    var formData = $('#frmTemporalRegistrar').serialize();
    //console.log(formData);


    $.ajax({
        url: "/Temporal/" + _accion,
        type: "POST",
        data: formData,
        contentType: 'application/x-www-form-urlencoded; charset=UTF-8', // optional
        success: function (data) {

            //console.log(data);

            if (data.operacion == true) {
                toastr.success(data.mensaje);

                loadDataTemporal();
                $("#modal-sm-temporal").modal('hide');
            }
            else {
                toastr.error(data.mensaje);
                //alert("Ocurrio un error!!");
            }

            if (data.operacion2 == false) {
                toastr.error(data.mensaje2);
            }

        },
        error: function () {
            toastr.error("Ocurrio un error!!");
            //alert("Ocurrio un error!!");
        }
    });

}


function GuardarDataPoliza() {
    var _nombreid = "#customFile";
    var _moneda = $("#v_moneda_poliza").val();
    var fileUpload = $(_nombreid).get(0);
    var files = fileUpload.files;


    if ($.trim(_moneda) == '') {
        toastr.error("Seleccione moneda");
        return false;
    }

    $("#modal-sm-polizas").modal('hide');
    $('#modal-loading').modal('show');

    // Creando objeto FormData
    var fileData = new FormData();

    // Looping over all files and add it to FormData object  
    for (var i = 0; i < files.length; i++) {
        fileData.append(files[i].name, files[i]);
        //alert(files[i].name);
    }

    // Adding one more key to FormData object  
    fileData.append('moneda', _moneda);
    fileData.append('anio', $("#v_anio_polizas").val());
    fileData.append('mes', $("#v_mes_polizas").val());

    $.ajax({
        url: '/Temporal/ExcelToDatabase',
        type: "POST",
        contentType: false, // Not to set any content header  
        processData: false, // Not to process data  
        data: fileData,
        success: function (data) {
            console.log(data);
            if (data.operacion == true) {
                toastr.success(data.mensaje);

                loadDataTemporal();
                $("#modal-sm-polizas").modal('hide');
                $('#modal-loading').modal('hide');
            }
            else {
                toastr.error(data.mensaje);
                //alert("Ocurrio un error!!");
            }
        },
        error: function (err) {
            alert_error(err.statusText);
        }
    });

}


function GuardarReservaVT() {
    var _moneda = $("#v_moneda_reserva_vt").val();

    if ($.trim(_moneda) == '') {
        toastr.error("Seleccione moneda");
        return false;
    }

    $("#modal-sm-procesar_vt").modal('hide');
    $('#modal-loading').modal('show');

    var formData = $('#frmProcesaReservaMatematicaVT').serialize();
    //console.log(formData);
    //alert("123");
    $.ajax({
        url: "/Temporal/ProcesaVT",
        type: "POST",
        data: formData,
        contentType: 'application/x-www-form-urlencoded; charset=UTF-8', // optional
        success: function (data) {

            //console.log(data);

            if (data.operacion == true) {
                toastr.success(data.mensaje);

                loadDataTemporal();
                $("#modal-sm-procesar_vt").modal('hide');
            }
            else {
                toastr.error(data.mensaje);
                //alert("Ocurrio un error!!");
            }

            if (data.operacion2 == false) {
                toastr.error(data.mensaje2);
            }

            $('#modal-loading').modal('hide');

        },
        error: function () {
            toastr.error("Ocurrio un error!!");
            $('#modal-loading').modal('hide');
            //alert("Ocurrio un error!!");
        }
    });
}


function IrResultadoVT(_anio, _mes) {
    const url = $("#RedirectResultadoVT").val();
    const new_url = url.replace("100_", _anio); /// modificar valores
    const new_url2 = new_url.replace("200_", _mes);
    //console.log("final  " + new_url2);
    location.href = new_url2;
}


function GuardarMoceVT() {
    var _moneda = $("#v_moneda_moce_vt").val();

    if ($.trim(_moneda) == '') {
        toastr.error("Seleccione moneda");
        return false;
    }

    $("#modal-sm-procesar_vt_moce").modal('hide');
    $('#modal-loading').modal('show');

    var formData = $('#frmProcesaMoceVT').serialize();
    //console.log(formData);
    //alert("123");
    $.ajax({
        url: "/Temporal/ProcesaMoceVT",
        type: "POST",
        data: formData,
        contentType: 'application/x-www-form-urlencoded; charset=UTF-8', // optional
        success: function (data) {

            //console.log(data);

            if (data.operacion == true) {
                toastr.success(data.mensaje);

                loadDataTemporal();
                $("#modal-sm-procesar_vt_moce").modal('hide');
            }
            else {
                toastr.error(data.mensaje);
                //alert("Ocurrio un error!!");
            }

            if (data.operacion2 == false) {
                toastr.error(data.mensaje2);
            }

            $('#modal-loading').modal('hide');

        },
        error: function () {
            toastr.error("Ocurrio un error!!");
            $('#modal-loading').modal('hide');
            //alert("Ocurrio un error!!");
        }
    });
}

