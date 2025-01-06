
/*
function mostrarresultvt() {
    $('#modal-sm-result-vt').modal('show');
}
*/

var __filas = 5000;

loadDataResultVT();

function loadDataResultVT() {
    var totregistros = 0;
    var paginas = 0;
    var _sw = 0;

    //alert("jim");
    var _moneda = $('#v_moneda2').val();
    var _anio = $('#v_anio2').val();
    var _mes = $('#v_mes2').val();

    if (_moneda == '') {
        _moneda = 'S';
    }

    if (_anio == 0 || _anio == '') {
        var _fecha = new Date();
        _anio = _fecha.getFullYear();
        _mes = _fecha.getMonth();
    }

    $.ajax({
        url: "/ResultadoVT/ListarResultado",
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

            $('#tabla_objeto_rvt').html("");
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

    $('#tabla_objeto_rvt')
        .append(
            $('<tr/>')
                //.append($('<td/>').html(_opt))
                //// item.NumeroPoliza
                .append($('<td/>').html(item.NumeroPoliza))
                .append($('<td/>').html(item.EstadoPoliza))
                .append($('<td/>').html(item.TasaUtilizada))
                .append($('<td/>').html(item.TasaVenta))
                .append($('<td/>').html(item.TLRA))
                .append($('<td/>').html(item.NPV_TV))
                .append($('<td/>').html(item.NPV_TLRA))
                .append($('<td/>').html(item.NPV_VOLA))
                .append($('<td/>').html(item.Check_TIR1))
                .append($('<td/>').html(item.Check_TIR2))
                .append($('<td/>').html(item.TasaMinima))
                .append($('<td/>').html(item.ReservaMinimaBruta))
                .append($('<td/>').html(item.ReservaMinimaCedida))
                .append($('<td/>').html(item.ReservaMinimaRetenida))
        );

}


//function mostrarresultvp(valor_excel, resultado, texto) {

//    $('#titulo_data').html(texto);
//    $('#v_excel').val(valor_excel);
//    $('#v_calculo').val(resultado);

//    $('#modal-sm-result-vp').modal('show');
//}


function export_excel_vt() {

    var _moneda = $("#v_moneda2").val();
    var _anio = $("#v_anio2").val();
    var _mes = $("#v_mes2").val();
    var _tipo = $("#v_export").val();
    var _poliza = $("#v_num_poliza").val();

    var _metodo = '';
    if (_tipo == 'R')
        _metodo = 'ExportarResultadoReserva';
    else {
        _metodo = 'ExportarDetalleReserva';

        if ($.trim(_poliza) == '') {
            toastr.error("Ingrese N° Poliza");
            return false;
        }
    }


    $.ajax({

        cache: false,
        url: '/ResultadoVT/' + _metodo,
        data: {
            anio: _anio,
            mes: _mes,
            moneda: _moneda,
            tipo: _tipo,
            poliza: _poliza
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