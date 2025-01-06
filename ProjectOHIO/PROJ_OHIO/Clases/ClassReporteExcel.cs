using System;
using System.Data;
using ClosedXML.Excel;
using System.IO;

namespace PROJ_OHIO.Clases
{
    public class ClassReporteExcel
    {

        public byte[] ExportarBusquedaDocumentosGeneral(DataTable resultado)
        {
            MemoryStream ms = new MemoryStream();

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sample Sheet");

                //// CABECERA
                worksheet.Cell("A1").Value = "Nro. Póliza";
                worksheet.Cell("B1").Value = "Estado de Póliza";

                worksheet.Cell("C1").Value = "Valor Póliza BOP";
                worksheet.Cell("D1").Value = "Valor Póliza BOP_Excel";
                worksheet.Cell("E1").Value = "Diferencia";

                worksheet.Cell("F1").Value = "Prima Pagada";
                worksheet.Cell("G1").Value = "Prima Pagada_Excel";
                worksheet.Cell("H1").Value = "Diferencia";

                ////worksheet.Cell("F1").Value = "Docu.Solic.";

                worksheet.Cell("I1").Value = "Interes";
                worksheet.Cell("J1").Value = "Interes_Excel";
                worksheet.Cell("K1").Value = "Diferencia";

                worksheet.Cell("L1").Value = "Rescates Parciales";
                worksheet.Cell("M1").Value = "Rescates Parciales_Excel";
                worksheet.Cell("N1").Value = "Diferencia";

                worksheet.Cell("O1").Value = "Cargo por Rescate Parcial";
                worksheet.Cell("P1").Value = "Cargo por Rescate Parcial_Excel";
                worksheet.Cell("Q1").Value = "Diferencia";

                worksheet.Cell("R1").Value = "Gasto Adm. PB";
                worksheet.Cell("S1").Value = "Gasto Adm. PB_Excel";
                worksheet.Cell("T1").Value = "Diferencia";

                worksheet.Cell("U1").Value = "Gasto Adm. Pexc";
                worksheet.Cell("V1").Value = "Gasto Adm. Pexc_Excel";
                worksheet.Cell("W1").Value = "Diferencia";

                worksheet.Cell("X1").Value = "Gasto Adm. Disminucion SA";
                worksheet.Cell("Y1").Value = "Gasto Adm. Disminucion SA_Excel";
                worksheet.Cell("Z1").Value = "Diferencia";

                worksheet.Cell("AA1").Value = "COI Muerte";
                worksheet.Cell("AB1").Value = "COI Muerte_Excel";
                worksheet.Cell("AC1").Value = "Diferencia";

                worksheet.Cell("AD1").Value = "COI ITP";
                worksheet.Cell("AE1").Value = "COI ITP_Excel";
                worksheet.Cell("AF1").Value = "Diferencia";

                worksheet.Cell("AG1").Value = "COI Muerte Accidental";
                worksheet.Cell("AH1").Value = "COI Muerte Accidental_Excel";
                worksheet.Cell("AI1").Value = "Diferencia";

                worksheet.Cell("AJ1").Value = "COI Enfermedades Graves";
                worksheet.Cell("AK1").Value = "COI Enfermedades Graves_Excel";
                worksheet.Cell("AL1").Value = "Diferencia";

                //worksheet.Cell("Z1").Value = "COI Enfermedades Graves";
                //worksheet.Cell("AA1").Value = "Excel";

                worksheet.Cell("AM1").Value = "COI Exoneracion";
                worksheet.Cell("AN1").Value = "COI Exoneracion_Excel";
                worksheet.Cell("AO1").Value = "Diferencia";

                worksheet.Cell("AP1").Value = "COI Muerte AA";
                worksheet.Cell("AQ1").Value = "COI Muerte AA_Excel";
                worksheet.Cell("AR1").Value = "Diferencia";

                //worksheet.Cell("AF1").Value = "COI Muerte AA";
                //worksheet.Cell("AG1").Value = "Excel";

                worksheet.Cell("AS1").Value = "COI ITP AA";
                worksheet.Cell("AT1").Value = "COI ITP AA_Excel";
                worksheet.Cell("AU1").Value = "Diferencia";

                worksheet.Cell("AV1").Value = "Cargo por Rescate Total Cancelación";
                worksheet.Cell("AW1").Value = "Cargo por Rescate Total Cancelación_Excel";
                worksheet.Cell("AX1").Value = "Diferencia";

                worksheet.Cell("AY1").Value = "Cois Pendientes Pago";
                worksheet.Cell("AZ1").Value = "Cois Pendientes Pago_Excel";
                worksheet.Cell("BA1").Value = "Diferencia";

                worksheet.Cell("BB1").Value = "Valor Poliza EOP";
                worksheet.Cell("BC1").Value = "Valor Poliza EOP_Excel";
                worksheet.Cell("BD1").Value = "Diferencia";

                worksheet.Cell("BE1").Value = "Valor Rescate";
                worksheet.Cell("BF1").Value = "Valor Rescate_Excel";
                worksheet.Cell("BG1").Value = "Diferencia";

                worksheet.Cell("BH1").Value = "Situación";

                ////worksheet.Cell("A1:J1").Style.Font.FontColor = XLColor.Red;
                worksheet.Range("A1:BH1").Style.Fill.BackgroundColor = XLColor.Gray;
                worksheet.Range("A1:BH1").Style.Font.Bold = true;
                worksheet.Range("A1:BH1").Style.Font.FontColor = XLColor.White;

                int i = 1;
                int flag;
                foreach (DataRow row in resultado.Rows)
                {
                    flag = 0;

                    worksheet.Cell("A" + (1 + i)).Value = Convert.ToString(row["NumeroPoliza"]);

                    worksheet.Cell("B" + (1 + i)).Value = Convert.ToString(row["EstadoPoliza"]);

                    worksheet.Cell("C" + (1 + i)).Value = Convert.ToDecimal(row["ValorPolizaBOP_2"]);
                    worksheet.Cell("D" + (1 + i)).Value = Convert.ToDecimal(row["ValorPolizaBOP_1"]);
                    worksheet.Cell("E" + (1 + i)).Value = Convert.ToDecimal(row["ValorPolizaBOP"]);
                    if (Convert.ToDecimal(row["ValorPolizaBOP"]) != 0)
                    {
                        // Valor Poliza BOP no se verificara como error
                        //flag = 1;
                        worksheet.Cell("E" + (1 + i)).Style.Font.FontColor = XLColor.Red;
                    }
                        

                    worksheet.Cell("F" + (1 + i)).Value = Convert.ToDecimal(row["PrimaPagada_2"]);
                    worksheet.Cell("G" + (1 + i)).Value = Convert.ToDecimal(row["PrimaPagada_1"]);
                    worksheet.Cell("H" + (1 + i)).Value = Convert.ToDecimal(row["PrimaPagada"]);
                    if (Convert.ToDecimal(row["PrimaPagada"]) != 0)
                    {
                        flag = 1;
                        worksheet.Cell("H" + (1 + i)).Style.Font.FontColor = XLColor.Red;
                    }
                    ////worksheet.Cell("F1").Value = "Docu.Solic.";

                    worksheet.Cell("I" + (1 + i)).Value = Convert.ToDecimal(row["Interes_2"]);
                    worksheet.Cell("J" + (1 + i)).Value = Convert.ToDecimal(row["Interes_1"]);
                    worksheet.Cell("K" + (1 + i)).Value = Convert.ToDecimal(row["Interes"]);
                    if (Convert.ToDecimal(row["Interes"]) != 0)
                    {
                        flag = 1;
                        worksheet.Cell("K" + (1 + i)).Style.Font.FontColor = XLColor.Red;
                    }

                    worksheet.Cell("L" + (1 + i)).Value = Convert.ToDecimal(row["RescatesParciales_2"]);
                    worksheet.Cell("M" + (1 + i)).Value = Convert.ToDecimal(row["RescatesParciales_1"]);
                    worksheet.Cell("N" + (1 + i)).Value = Convert.ToDecimal(row["RescatesParciales"]);
                    if (Convert.ToDecimal(row["RescatesParciales"]) != 0)
                    {
                        flag = 1;
                        worksheet.Cell("N" + (1 + i)).Style.Font.FontColor = XLColor.Red;
                    }

                    worksheet.Cell("O" + (1 + i)).Value = Convert.ToDecimal(row["CargoPorRescateParcial_2"]);
                    worksheet.Cell("P" + (1 + i)).Value = Convert.ToDecimal(row["CargoPorRescateParcial_1"]);
                    worksheet.Cell("Q" + (1 + i)).Value = Convert.ToDecimal(row["CargoPorRescateParcial"]);
                    if (Convert.ToDecimal(row["CargoPorRescateParcial"]) != 0)
                    {
                        flag = 1;
                        worksheet.Cell("Q" + (1 + i)).Style.Font.FontColor = XLColor.Red;
                    }

                    worksheet.Cell("R" + (1 + i)).Value = Convert.ToDecimal(row["GastoAdministrativoPB_2"]);
                    worksheet.Cell("S" + (1 + i)).Value = Convert.ToDecimal(row["GastoAdministrativoPB_1"]);
                    worksheet.Cell("T" + (1 + i)).Value = Convert.ToDecimal(row["GastoAdministrativoPB"]);
                    if (Convert.ToDecimal(row["GastoAdministrativoPB"]) != 0)
                    {
                        flag = 1;
                        worksheet.Cell("T" + (1 + i)).Style.Font.FontColor = XLColor.Red;
                    }

                    worksheet.Cell("U" + (1 + i)).Value = Convert.ToDecimal(row["GastoAdministrativoPexc_2"]);
                    worksheet.Cell("V" + (1 + i)).Value = Convert.ToDecimal(row["GastoAdministrativoPexc_1"]);
                    worksheet.Cell("W" + (1 + i)).Value = Convert.ToDecimal(row["GastoAdministrativoPexc"]);
                    if (Convert.ToDecimal(row["GastoAdministrativoPexc"]) != 0)
                    {
                        flag = 1;
                        worksheet.Cell("W" + (1 + i)).Style.Font.FontColor = XLColor.Red;
                    }

                    worksheet.Cell("X" + (1 + i)).Value = Convert.ToDecimal(row["GastoAdministrativoDisminucionSA_2"]);
                    worksheet.Cell("Y" + (1 + i)).Value = Convert.ToDecimal(row["GastoAdministrativoDisminucionSA_1"]);
                    worksheet.Cell("Z" + (1 + i)).Value = Convert.ToDecimal(row["GastoAdministrativoDisminucionSA"]);
                    if (Convert.ToDecimal(row["GastoAdministrativoDisminucionSA"]) != 0)
                    {
                        flag = 1;
                        worksheet.Cell("Z" + (1 + i)).Style.Font.FontColor = XLColor.Red;
                    }

                    worksheet.Cell("AA" + (1 + i)).Value = Convert.ToDecimal(row["COI_Muerte_2"]);
                    worksheet.Cell("AB" + (1 + i)).Value = Convert.ToDecimal(row["COI_Muerte_1"]);
                    worksheet.Cell("AC" + (1 + i)).Value = Convert.ToDecimal(row["COI_Muerte"]);
                    if (Convert.ToDecimal(row["COI_Muerte"]) != 0)
                    {
                        flag = 1;
                        worksheet.Cell("AC" + (1 + i)).Style.Font.FontColor = XLColor.Red;
                    }

                    worksheet.Cell("AD" + (1 + i)).Value = Convert.ToDecimal(row["COI_ITP_2"]);
                    worksheet.Cell("AE" + (1 + i)).Value = Convert.ToDecimal(row["COI_ITP_1"]);
                    worksheet.Cell("AF" + (1 + i)).Value = Convert.ToDecimal(row["COI_ITP"]);
                    if (Convert.ToDecimal(row["COI_ITP"]) != 0)
                    {
                        flag = 1;
                        worksheet.Cell("AF" + (1 + i)).Style.Font.FontColor = XLColor.Red;
                    }

                    worksheet.Cell("AG" + (1 + i)).Value = Convert.ToDecimal(row["COI_MuerteAccidental_2"]);
                    worksheet.Cell("AH" + (1 + i)).Value = Convert.ToDecimal(row["COI_MuerteAccidental_1"]);
                    worksheet.Cell("AI" + (1 + i)).Value = Convert.ToDecimal(row["COI_MuerteAccidental"]);
                    if (Convert.ToDecimal(row["COI_MuerteAccidental"]) != 0)
                    {
                        flag = 1;
                        worksheet.Cell("AI" + (1 + i)).Style.Font.FontColor = XLColor.Red;
                    }

                    worksheet.Cell("AJ" + (1 + i)).Value = Convert.ToDecimal(row["COI_EnfermedadesGraves_2"]);
                    worksheet.Cell("AK" + (1 + i)).Value = Convert.ToDecimal(row["COI_EnfermedadesGraves_1"]);
                    worksheet.Cell("AL" + (1 + i)).Value = Convert.ToDecimal(row["COI_EnfermedadesGraves"]);
                    if (Convert.ToDecimal(row["COI_EnfermedadesGraves"]) != 0)
                    {
                        flag = 1;
                        worksheet.Cell("AL" + (1 + i)).Style.Font.FontColor = XLColor.Red;
                    }

                    //worksheet.Cell("Z" + (1 + i)).Value = "COI Enfermedades Graves";
                    //worksheet.Cell("AA" + (1 + i)).Value = "Excel";

                    worksheet.Cell("AM" + (1 + i)).Value = Convert.ToDecimal(row["COI_Exoneracion_2"]);
                    worksheet.Cell("AN" + (1 + i)).Value = Convert.ToDecimal(row["COI_Exoneracion_1"]);
                    worksheet.Cell("AO" + (1 + i)).Value = Convert.ToDecimal(row["COI_Exoneracion"]);
                    if (Convert.ToDecimal(row["COI_Exoneracion"]) != 0)
                    {
                        flag = 1;
                        worksheet.Cell("AO" + (1 + i)).Style.Font.FontColor = XLColor.Red;
                    }

                    worksheet.Cell("AP" + (1 + i)).Value = Convert.ToDecimal(row["COI_MuerteAA_2"]);
                    worksheet.Cell("AQ" + (1 + i)).Value = Convert.ToDecimal(row["COI_MuerteAA_1"]);
                    worksheet.Cell("AR" + (1 + i)).Value = Convert.ToDecimal(row["COI_MuerteAA"]);
                    if (Convert.ToDecimal(row["COI_MuerteAA"]) != 0)
                    {
                        flag = 1;
                        worksheet.Cell("AR" + (1 + i)).Style.Font.FontColor = XLColor.Red;
                    }

                    //worksheet.Cell("AF" + (1 + i)).Value = "COI Muerte AA";
                    //worksheet.Cell("AG" + (1 + i)).Value = "Excel";

                    worksheet.Cell("AS" + (1 + i)).Value = Convert.ToDecimal(row["COI_ITP_AA_2"]);
                    worksheet.Cell("AT" + (1 + i)).Value = Convert.ToDecimal(row["COI_ITP_AA_1"]);
                    worksheet.Cell("AU" + (1 + i)).Value = Convert.ToDecimal(row["COI_ITP_AA"]);
                    if (Convert.ToDecimal(row["COI_ITP_AA"]) != 0)
                    {
                        flag = 1;
                        worksheet.Cell("AU" + (1 + i)).Style.Font.FontColor = XLColor.Red;
                    }

                    worksheet.Cell("AV" + (1 + i)).Value = Convert.ToDecimal(row["CargoPorRescateTotalCancelacion_2"]);
                    worksheet.Cell("AW" + (1 + i)).Value = Convert.ToDecimal(row["CargoPorRescateTotalCancelacion_1"]);
                    worksheet.Cell("AX" + (1 + i)).Value = Convert.ToDecimal(row["CargoPorRescateTotalCancelacion"]);
                    if (Convert.ToDecimal(row["CargoPorRescateTotalCancelacion"]) != 0)
                    {
                        flag = 1;
                        worksheet.Cell("AX" + (1 + i)).Style.Font.FontColor = XLColor.Red;
                    }

                    worksheet.Cell("AY" + (1 + i)).Value = Convert.ToDecimal(row["COIS_PendientesPago_2"]);
                    worksheet.Cell("AZ" + (1 + i)).Value = Convert.ToDecimal(row["COIS_PendientesPago_1"]);
                    worksheet.Cell("BA" + (1 + i)).Value = Convert.ToDecimal(row["COIS_PendientesPago"]);
                    if (Convert.ToDecimal(row["COIS_PendientesPago"]) != 0)
                    {
                        flag = 1;
                        worksheet.Cell("BA" + (1 + i)).Style.Font.FontColor = XLColor.Red;
                    }

                    worksheet.Cell("BB" + (1 + i)).Value = Convert.ToDecimal(row["ValorPolizaEOP_2"]);
                    worksheet.Cell("BC" + (1 + i)).Value = Convert.ToDecimal(row["ValorPolizaEOP_1"]);
                    worksheet.Cell("BD" + (1 + i)).Value = Convert.ToDecimal(row["ValorPolizaEOP"]);
                    if (Convert.ToDecimal(row["ValorPolizaEOP"]) != 0)
                    {
                        flag = 1;
                        worksheet.Cell("BD" + (1 + i)).Style.Font.FontColor = XLColor.Red;
                    }

                    worksheet.Cell("BE" + (1 + i)).Value = Convert.ToDecimal(row["ValorRescate_2"]);
                    worksheet.Cell("BF" + (1 + i)).Value = Convert.ToDecimal(row["ValorRescate_1"]);
                    worksheet.Cell("BG" + (1 + i)).Value = Convert.ToDecimal(row["ValorRescate"]);
                    if (Convert.ToDecimal(row["ValorRescate"]) != 0)
                    {
                        flag = 1;
                        worksheet.Cell("BG" + (1 + i)).Style.Font.FontColor = XLColor.Red;
                    }

                    if (flag == 1)
                    {
                        worksheet.Cell("BH" + (1 + i)).Value = "ERROR";
                        worksheet.Cell("BH" + (1 + i)).Style.Font.FontColor = XLColor.Red;
                    }
                    else
                    {
                        worksheet.Cell("BH" + (1 + i)).Value = "";
                    }
                    

                    i = i + 1;
                }

                worksheet.ColumnsUsed().AdjustToContents();
                worksheet.Name = "Datos";

                using (MemoryStream stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    stream.Position = 0;
                    ms = stream;
                }
            }

            return ms.ToArray();
        }

        public byte[] ExportarResultadoReservaMatematicaVU(DataTable resultado)
        {
            MemoryStream ms = new MemoryStream();

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sample Sheet");

                //// CABECERA
                worksheet.Cell("A1").Value = "Nro. Póliza";
                worksheet.Cell("B1").Value = "Estado de Póliza";

                worksheet.Cell("C1").Value = "TasaUtilizada";
                worksheet.Cell("D1").Value = "TasaVenta";
                worksheet.Cell("E1").Value = "TLRA";

                worksheet.Cell("F1").Value = "NPV_TV";
                worksheet.Cell("G1").Value = "NPV_TLRA";
                worksheet.Cell("H1").Value = "NPV_VOLA";

                ////worksheet.Cell("F1").Value = "Docu.Solic.";

                worksheet.Cell("I1").Value = "Check_TIR1";
                worksheet.Cell("J1").Value = "Check_TIR2";
                worksheet.Cell("K1").Value = "TasaMinima";

                worksheet.Cell("L1").Value = "ReservaMinimaBruta";
                worksheet.Cell("M1").Value = "ReservaMinimaCedida";
                worksheet.Cell("N1").Value = "ReservaMinimaRetenida";

                ////worksheet.Cell("A1:J1").Style.Font.FontColor = XLColor.Red;
                worksheet.Range("A1:N1").Style.Fill.BackgroundColor = XLColor.Gray;
                worksheet.Range("A1:N1").Style.Font.Bold = true;
                worksheet.Range("A1:N1").Style.Font.FontColor = XLColor.White;

                int i = 1;
                foreach (DataRow row in resultado.Rows)
                {

                    worksheet.Cell("A" + (1 + i)).Value = Convert.ToString(row["NumeroPoliza"]);

                    worksheet.Cell("B" + (1 + i)).Value = Convert.ToString(row["EstadoPoliza"]);

                    worksheet.Cell("C" + (1 + i)).Value = Convert.ToString(row["TasaUtilizada"]);
                    worksheet.Cell("D" + (1 + i)).Value = Convert.ToDecimal(row["TasaVenta"]);
                    worksheet.Cell("E" + (1 + i)).Value = Convert.ToDecimal(row["TLRA"]);


                    worksheet.Cell("F" + (1 + i)).Value = Convert.ToDecimal(row["NPV_TV"]);
                    worksheet.Cell("G" + (1 + i)).Value = Convert.ToDecimal(row["NPV_TLRA"]);
                    worksheet.Cell("H" + (1 + i)).Value = Convert.ToDecimal(row["NPV_VOLA"]);



                    worksheet.Cell("I" + (1 + i)).Value = Convert.ToDecimal(row["Check_TIR1"]);
                    worksheet.Cell("J" + (1 + i)).Value = Convert.ToDecimal(row["Check_TIR2"]);
                    worksheet.Cell("K" + (1 + i)).Value = Convert.ToDecimal(row["TasaMinima"]);

                    worksheet.Cell("L" + (1 + i)).Value = Convert.ToDecimal(row["ReservaMinimaBruta"]);
                    worksheet.Cell("M" + (1 + i)).Value = Convert.ToDecimal(row["ReservaMinimaCedida"]);
                    worksheet.Cell("N" + (1 + i)).Value = Convert.ToDecimal(row["ReservaMinimaRetenida"]);

                    i = i + 1;
                }

                worksheet.ColumnsUsed().AdjustToContents();
                worksheet.Name = "Datos";

                using (MemoryStream stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    stream.Position = 0;
                    ms = stream;
                }
            }

            return ms.ToArray();
        }

        public byte[] ExportarDetalleReservaMatematicaVU(DataTable resultado)
        {
            MemoryStream ms = new MemoryStream();

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sample Sheet");

                //// CABECERA
                worksheet.Cell("A1").Value = "Nro. Póliza";
                worksheet.Cell("B1").Value = "Estado de Póliza";

                worksheet.Cell("C1").Value = "fecha_evaluacion";
                worksheet.Cell("D1").Value = "anio_poliza";
                worksheet.Cell("E1").Value = "mes_poliza";
                worksheet.Cell("F1").Value = "dia_pago";
                worksheet.Cell("G1").Value = "edad_real_titular";
                worksheet.Cell("H1").Value = "edad_actuarial_titular";
                worksheet.Cell("I1").Value = "edad_cois_titular";
                worksheet.Cell("J1").Value = "edad_real_adicional";
                worksheet.Cell("K1").Value = "edad_actuarial_adicional";
                worksheet.Cell("L1").Value = "edad_cois_adicional";
                worksheet.Cell("M1").Value = "cobertura_muerte";
                worksheet.Cell("N1").Value = "cobertura_itp";

                worksheet.Cell("O1").Value = "cobertura_muerte_accidental";
                worksheet.Cell("P1").Value = "cobertura_enfermedades_graves";
                worksheet.Cell("Q1").Value = "cobertura_muerte_aa";
                worksheet.Cell("R1").Value = "cobertura_itp_aa";
                worksheet.Cell("S1").Value = "cargo_admin_pb";
                worksheet.Cell("T1").Value = "cargo_admin_excess";
                worksheet.Cell("U1").Value = "tasa_interes_acreditado_mensual";
                worksheet.Cell("V1").Value = "valor_poliza_bop";
                worksheet.Cell("W1").Value = "valor_poliza_o";
                worksheet.Cell("X1").Value = "prima_pagada";
                worksheet.Cell("Y1").Value = "rescate_parcial";
                worksheet.Cell("Z1").Value = "cargo_rescate_parcial";

                worksheet.Cell("AA1").Value = "ffactor";
                worksheet.Cell("AB1").Value = "gasto_administrativo_prima_basica";
                worksheet.Cell("AC1").Value = "gasto_administrativo_prima_excedente";
                worksheet.Cell("AD1").Value = "gasto_admin_disminucion_sa";
                worksheet.Cell("AE1").Value = "interes";
                worksheet.Cell("AF1").Value = "suma_asegurada_muerte";
                worksheet.Cell("AG1").Value = "monto_riesgo_muerte_bop";
                worksheet.Cell("AH1").Value = "suma_asegurada_itp";
                worksheet.Cell("AI1").Value = "suma_asegurada_muerte_accidental";
                worksheet.Cell("AJ1").Value = "suma_asegurada_enfermedades_graves";
                worksheet.Cell("AK1").Value = "suma_asegurada_muerte_aa";
                worksheet.Cell("AL1").Value = "suma_asegurada_itp_aa";

                worksheet.Cell("AM1").Value = "coi_muerte";
                worksheet.Cell("AN1").Value = "coi_itp";
                worksheet.Cell("AO1").Value = "coi_muerte_accidental";
                worksheet.Cell("AP1").Value = "coi_enfermedades_graves";
                worksheet.Cell("AQ1").Value = "coi_muerte_aa";
                worksheet.Cell("AR1").Value = "coi_itp_aa";
                worksheet.Cell("AS1").Value = "coi_exoneracion";
                worksheet.Cell("AT1").Value = "monto_riesgo_muerte_eop";
                worksheet.Cell("AU1").Value = "cargo_rescate_total";
                worksheet.Cell("AV1").Value = "cois_pendientes_pago";
                worksheet.Cell("AW1").Value = "valor_poliza_eop";
                worksheet.Cell("AX1").Value = "cargo_por_rescate";

                worksheet.Cell("AY1").Value = "valor_rescate";
                worksheet.Cell("AZ1").Value = "edad_meses_titular";
                worksheet.Cell("BA1").Value = "edad_anios_t_titular";
                worksheet.Cell("BB1").Value = "edad_anios_t_mas_1_titular";
                worksheet.Cell("BC1").Value = "lx_t_titular";
                worksheet.Cell("BD1").Value = "lx_t_mas_1_titular";
                worksheet.Cell("BE1").Value = "lx_mortalidad_probabilidad_titular";
                worksheet.Cell("BF1").Value = "edad_meses_adicional";
                worksheet.Cell("BG1").Value = "edad_anios_t_adicional";
                worksheet.Cell("BH1").Value = "edad_anios_t_mas_1_adicional";
                worksheet.Cell("BI1").Value = "lx_t_adicional";
                worksheet.Cell("BJ1").Value = "lx_t_mas_1_adicional";

                worksheet.Cell("BK1").Value = "lx_mortalidad_probabilidad_adicional";
                worksheet.Cell("BL1").Value = "edad_meses_titular_cat";
                worksheet.Cell("BM1").Value = "edad_anios_t_titular_cat";
                worksheet.Cell("BN1").Value = "edad_anios_t_mas_1_titular_cat";
                worksheet.Cell("BO1").Value = "lx_t_titular_cat";
                worksheet.Cell("BP1").Value = "lx_t_mas_1_titular_cat";
                worksheet.Cell("BQ1").Value = "lx_catastrofe_probabilidad_titular_cat";
                worksheet.Cell("BR1").Value = "edad_meses_adicional_cat";
                worksheet.Cell("BS1").Value = "edad_anios_t_adicional_cat";
                worksheet.Cell("BT1").Value = "edad_anios_t_mas_1_adicional_cat";
                worksheet.Cell("BU1").Value = "lx_t_adicional_cat";
                worksheet.Cell("BV1").Value = "lx_t_mas_1_adicional_cat";

                worksheet.Cell("BW1").Value = "lx_catastrofe_probabilidad_adicional_cat";
                worksheet.Cell("BX1").Value = "suma_asegurada_muerte_exoneracion";
                worksheet.Cell("BY1").Value = "qx_tipo_riesgo_mortalidad_titular";
                worksheet.Cell("BZ1").Value = "qx_tipo_riesgo_mortalidad_adicional";
                worksheet.Cell("CA1").Value = "qx_tipo_riesgo_catastrofe_titular";
                worksheet.Cell("CB1").Value = "qx_tipo_riesgo_catastrofe_adicional";
                worksheet.Cell("CC1").Value = "t";
                worksheet.Cell("CD1").Value = "tpx_inicial";
                worksheet.Cell("CE1").Value = "ajd_qx_monthly";
                worksheet.Cell("CF1").Value = "lapses";
                worksheet.Cell("CG1").Value = "invalidez";
                worksheet.Cell("CH1").Value = "muerte_accidental";

                worksheet.Cell("CI1").Value = "enfermedades_graves";
                worksheet.Cell("CJ1").Value = "exoneracion";
                worksheet.Cell("CK1").Value = "muerte_asegurado_adicional";
                worksheet.Cell("CL1").Value = "itp_asegurado_adicional";
                worksheet.Cell("CM1").Value = "tpx_final";
                worksheet.Cell("CN1").Value = "aux_uno_menos_invalidez";
                worksheet.Cell("CO1").Value = "aux_uno_menos_enfermedades_graves";
                worksheet.Cell("CP1").Value = "aux_uno_menos_muerte_aa";
                worksheet.Cell("CQ1").Value = "aux_uno_menos_itp_aa";
                worksheet.Cell("CR1").Value = "sx_muerte";
                worksheet.Cell("CS1").Value = "sx_muerte_accidental";
                worksheet.Cell("CT1").Value = "sx_invalidez";

                worksheet.Cell("CU1").Value = "sx_enfermedades_graves";
                worksheet.Cell("CV1").Value = "sx_exoneracion";
                worksheet.Cell("CW1").Value = "sx_muerte_aa";
                worksheet.Cell("CX1").Value = "sx_itp_aa";
                worksheet.Cell("CY1").Value = "prima_comercial";
                worksheet.Cell("CZ1").Value = "interes_obtenido";
                worksheet.Cell("DA1").Value = "frmb_cargo_rescate_total";
                worksheet.Cell("DB1").Value = "frmb_eg_muerte";
                worksheet.Cell("DC1").Value = "frmb_eg_itp";
                worksheet.Cell("DD1").Value = "frmb_eg_ma";
                worksheet.Cell("DE1").Value = "frmb_eg_exo";
                worksheet.Cell("DF1").Value = "frmb_eg_enf_graves";

                worksheet.Cell("DG1").Value = "frmb_eg_muerte_aa";
                worksheet.Cell("DH1").Value = "frmb_eg_itp_aa";
                worksheet.Cell("DI1").Value = "frmb_eg_valor_poliza";
                worksheet.Cell("DJ1").Value = "frmb_eg_comision_basica";
                worksheet.Cell("DK1").Value = "frmb_eg_comision_excedente";
                worksheet.Cell("DL1").Value = "frmb_eg_gastos_administrativos";
                worksheet.Cell("DM1").Value = "frmb_eg_g_gestion_siniestros";
                worksheet.Cell("DN1").Value = "frmb_eg_gastos_inversiones";
                worksheet.Cell("DO1").Value = "mv_periodo";
                worksheet.Cell("DP1").Value = "mv_vtd_vola_mensual";
                worksheet.Cell("DQ1").Value = "mv_uno_uno_i_t";
                worksheet.Cell("DR1").Value = "mv_flujo_bop";

                worksheet.Cell("DS1").Value = "mv_flujo_eop";
                worksheet.Cell("DT1").Value = "mv_flujo_bel_descontado";
                worksheet.Cell("DU1").Value = "mtv_periodo";
                worksheet.Cell("DV1").Value = "mtv_tasa_venta_mensual";
                worksheet.Cell("DW1").Value = "mtv_uno_uno_i_t";
                worksheet.Cell("DX1").Value = "mtv_flujo_bop";
                worksheet.Cell("DY1").Value = "mtv_flujo_eop";
                worksheet.Cell("DZ1").Value = "mtv_flujo_bel_descontado";
                worksheet.Cell("EA1").Value = "mtlra_periodo";
                worksheet.Cell("EB1").Value = "mtlra_tir";
                worksheet.Cell("EC1").Value = "mtlra_tlra";
                worksheet.Cell("ED1").Value = "mtlra_flujo_bop";

                worksheet.Cell("ED1").Value = "mtlra_flujo_eop";
                worksheet.Cell("EE1").Value = "mtlra_flujo_bel_descontado";
                worksheet.Cell("EF1").Value = "mtm_periodo";
                worksheet.Cell("EG1").Value = "mtm_tasa_minima";
                worksheet.Cell("EH1").Value = "mtm_fd_tasa_minima";
                worksheet.Cell("EI1").Value = "mtm_flujo_bop";
                worksheet.Cell("EJ1").Value = "mtm_flujo_eop";
                worksheet.Cell("EK1").Value = "mtm_flujo_bel_descontado";
                worksheet.Cell("EL1").Value = "pc_muerte";
                worksheet.Cell("EM1").Value = "pc_itp";
                worksheet.Cell("EN1").Value = "pc_ma";
                worksheet.Cell("EO1").Value = "pc_muerte_aa";

                worksheet.Cell("EP1").Value = "pc_itp_aa";
                worksheet.Cell("EQ1").Value = "ci_muerte";
                worksheet.Cell("ER1").Value = "ci_itp";
                worksheet.Cell("ES1").Value = "ci_ma";
                worksheet.Cell("ET1").Value = "ci_muerte_aa";
                worksheet.Cell("EU1").Value = "ci_itp_aa";
                worksheet.Cell("EV1").Value = "ce_muerte";
                worksheet.Cell("EW1").Value = "ce_itp";
                worksheet.Cell("EX1").Value = "ce_ma";
                worksheet.Cell("EY1").Value = "ce_muerte_aa";
                worksheet.Cell("EZ1").Value = "ce_itp_aa";
                worksheet.Cell("FA1").Value = "ce_menos_ci";

                worksheet.Cell("FB1").Value = "ce_menos_ci_tasa_minima";

               
                ////worksheet.Cell("A1:J1").Style.Font.FontColor = XLColor.Red;
                worksheet.Range("A1:FB1").Style.Fill.BackgroundColor = XLColor.Gray;
                worksheet.Range("A1:FB1").Style.Font.Bold = true;
                worksheet.Range("A1:FB1").Style.Font.FontColor = XLColor.White;

                int i = 1;
                foreach (DataRow row in resultado.Rows)
                {

                    worksheet.Cell("A" + (1 + i)).Value = Convert.ToString(row["NumeroPoliza"]);
                    worksheet.Cell("B" + (1 + i)).Value = Convert.ToString(row["EstadoPoliza"]);

                    worksheet.Cell("C" + (1 + i)).Value = Convert.ToString(row["fecha_evaluacion"]);
                    worksheet.Cell("D" + (1 + i)).Value = Convert.ToString(row["anio_poliza"]);
                    worksheet.Cell("E" + (1 + i)).Value = Convert.ToString(row["mes_poliza"]);
                    worksheet.Cell("F" + (1 + i)).Value = Convert.ToString(row["dia_pago"]);
                    worksheet.Cell("G" + (1 + i)).Value = Convert.ToString(row["edad_real_titular"]);
                    worksheet.Cell("H" + (1 + i)).Value = Convert.ToString(row["edad_actuarial_titular"]);
                    worksheet.Cell("I" + (1 + i)).Value = Convert.ToString(row["edad_cois_titular"]);
                    worksheet.Cell("J" + (1 + i)).Value = Convert.ToString(row["edad_real_adicional"]);
                    worksheet.Cell("K" + (1 + i)).Value = Convert.ToString(row["edad_actuarial_adicional"]);
                    worksheet.Cell("L" + (1 + i)).Value = Convert.ToString(row["edad_cois_adicional"]);

                    worksheet.Cell("M" + (1 + i)).Value = Convert.ToDecimal(row["cobertura_muerte"]);
                    worksheet.Cell("N" + (1 + i)).Value = Convert.ToDecimal(row["cobertura_itp"]);
                    worksheet.Cell("O" + (1 + i)).Value = Convert.ToDecimal(row["cobertura_muerte_accidental"]);
                    worksheet.Cell("P" + (1 + i)).Value = Convert.ToDecimal(row["cobertura_enfermedades_graves"]);
                    worksheet.Cell("Q" + (1 + i)).Value = Convert.ToDecimal(row["cobertura_muerte_aa"]);
                    worksheet.Cell("R" + (1 + i)).Value = Convert.ToDecimal(row["cobertura_itp_aa"]);
                    worksheet.Cell("S" + (1 + i)).Value = Convert.ToDecimal(row["cargo_admin_pb"]);
                    worksheet.Cell("T" + (1 + i)).Value = Convert.ToDecimal(row["cargo_admin_excess"]);
                    worksheet.Cell("U" + (1 + i)).Value = Convert.ToDecimal(row["tasa_interes_acreditado_mensual"]);
                    worksheet.Cell("V" + (1 + i)).Value = Convert.ToDecimal(row["valor_poliza_bop"]);
                    worksheet.Cell("W" + (1 + i)).Value = Convert.ToDecimal(row["valor_poliza_o"]);
                    worksheet.Cell("X" + (1 + i)).Value = Convert.ToDecimal(row["prima_pagada"]);
                    worksheet.Cell("Y" + (1 + i)).Value = Convert.ToDecimal(row["rescate_parcial"]);
                    worksheet.Cell("Z" + (1 + i)).Value = Convert.ToDecimal(row["cargo_rescate_parcial"]);
                    worksheet.Cell("AA" + (1 + i)).Value = Convert.ToDecimal(row["ffactor"]);
                    worksheet.Cell("AB" + (1 + i)).Value = Convert.ToDecimal(row["gasto_administrativo_prima_basica"]);
                    worksheet.Cell("AC" + (1 + i)).Value = Convert.ToDecimal(row["gasto_administrativo_prima_excedente"]);
                    worksheet.Cell("AD" + (1 + i)).Value = Convert.ToDecimal(row["gasto_admin_disminucion_sa"]);
                    worksheet.Cell("AE" + (1 + i)).Value = Convert.ToDecimal(row["interes"]);
                    worksheet.Cell("AF" + (1 + i)).Value = Convert.ToDecimal(row["suma_asegurada_muerte"]);
                    worksheet.Cell("AG" + (1 + i)).Value = Convert.ToDecimal(row["monto_riesgo_muerte_bop"]);
                    worksheet.Cell("AH" + (1 + i)).Value = Convert.ToDecimal(row["suma_asegurada_itp"]);
                    worksheet.Cell("AI" + (1 + i)).Value = Convert.ToDecimal(row["suma_asegurada_muerte_accidental"]);
                    worksheet.Cell("AJ" + (1 + i)).Value = Convert.ToDecimal(row["suma_asegurada_enfermedades_graves"]);
                    worksheet.Cell("AK" + (1 + i)).Value = Convert.ToDecimal(row["suma_asegurada_muerte_aa"]);
                    worksheet.Cell("AL" + (1 + i)).Value = Convert.ToDecimal(row["suma_asegurada_itp_aa"]);
                    worksheet.Cell("AM" + (1 + i)).Value = Convert.ToDecimal(row["coi_muerte"]);
                    worksheet.Cell("AN" + (1 + i)).Value = Convert.ToDecimal(row["coi_itp"]);
                    worksheet.Cell("AO" + (1 + i)).Value = Convert.ToDecimal(row["coi_muerte_accidental"]);
                    worksheet.Cell("AP" + (1 + i)).Value = Convert.ToDecimal(row["coi_enfermedades_graves"]);
                    worksheet.Cell("AQ" + (1 + i)).Value = Convert.ToDecimal(row["coi_muerte_aa"]);
                    worksheet.Cell("AR" + (1 + i)).Value = Convert.ToDecimal(row["coi_itp_aa"]);
                    worksheet.Cell("AS" + (1 + i)).Value = Convert.ToDecimal(row["coi_exoneracion"]);
                    worksheet.Cell("AT" + (1 + i)).Value = Convert.ToDecimal(row["monto_riesgo_muerte_eop"]);
                    worksheet.Cell("AU" + (1 + i)).Value = Convert.ToDecimal(row["cargo_rescate_total"]);
                    worksheet.Cell("AV" + (1 + i)).Value = Convert.ToDecimal(row["cois_pendientes_pago"]);
                    worksheet.Cell("AW" + (1 + i)).Value = Convert.ToDecimal(row["valor_poliza_eop"]);
                    worksheet.Cell("AX" + (1 + i)).Value = Convert.ToDecimal(row["cargo_por_rescate"]);
                    worksheet.Cell("AY" + (1 + i)).Value = Convert.ToDecimal(row["valor_rescate"]);

                    worksheet.Cell("AZ" + (1 + i)).Value = Convert.ToString(row["edad_meses_titular"]);
                    worksheet.Cell("BA" + (1 + i)).Value = Convert.ToString(row["edad_anios_t_titular"]);
                    worksheet.Cell("BB" + (1 + i)).Value = Convert.ToString(row["edad_anios_t_mas_1_titular"]);

                    worksheet.Cell("BC" + (1 + i)).Value = Convert.ToDecimal(row["lx_t_titular"]);
                    worksheet.Cell("BD" + (1 + i)).Value = Convert.ToDecimal(row["lx_t_mas_1_titular"]);
                    worksheet.Cell("BE" + (1 + i)).Value = Convert.ToDecimal(row["lx_mortalidad_probabilidad_titular"]);

                    worksheet.Cell("BF" + (1 + i)).Value = Convert.ToString(row["edad_meses_adicional"]);
                    worksheet.Cell("BG" + (1 + i)).Value = Convert.ToString(row["edad_anios_t_adicional"]);
                    worksheet.Cell("BH" + (1 + i)).Value = Convert.ToString(row["edad_anios_t_mas_1_adicional"]);

                    worksheet.Cell("BI" + (1 + i)).Value = Convert.ToDecimal(row["lx_t_adicional"]);
                    worksheet.Cell("BJ" + (1 + i)).Value = Convert.ToDecimal(row["lx_t_mas_1_adicional"]);
                    worksheet.Cell("BK" + (1 + i)).Value = Convert.ToDecimal(row["lx_mortalidad_probabilidad_adicional"]);

                    worksheet.Cell("BL" + (1 + i)).Value = Convert.ToString(row["edad_meses_titular_cat"]);
                    worksheet.Cell("BM" + (1 + i)).Value = Convert.ToString(row["edad_anios_t_titular_cat"]);
                    worksheet.Cell("BN" + (1 + i)).Value = Convert.ToString(row["edad_anios_t_mas_1_titular_cat"]);

                    worksheet.Cell("BO" + (1 + i)).Value = Convert.ToDecimal(row["lx_t_titular_cat"]);
                    worksheet.Cell("BP" + (1 + i)).Value = Convert.ToDecimal(row["lx_t_mas_1_titular_cat"]);
                    worksheet.Cell("BQ" + (1 + i)).Value = Convert.ToDecimal(row["lx_catastrofe_probabilidad_titular_cat"]);

                    worksheet.Cell("BR" + (1 + i)).Value = Convert.ToString(row["edad_meses_adicional_cat"]);
                    worksheet.Cell("BS" + (1 + i)).Value = Convert.ToString(row["edad_anios_t_adicional_cat"]);
                    worksheet.Cell("BT" + (1 + i)).Value = Convert.ToString(row["edad_anios_t_mas_1_adicional_cat"]);

                    worksheet.Cell("BU" + (1 + i)).Value = Convert.ToDecimal(row["lx_t_adicional_cat"]);
                    worksheet.Cell("BV" + (1 + i)).Value = Convert.ToDecimal(row["lx_t_mas_1_adicional_cat"]);
                    worksheet.Cell("BW" + (1 + i)).Value = Convert.ToDecimal(row["lx_catastrofe_probabilidad_adicional_cat"]);
                    worksheet.Cell("BX" + (1 + i)).Value = Convert.ToDecimal(row["suma_asegurada_muerte_exoneracion"]); //// error null
                    worksheet.Cell("BY" + (1 + i)).Value = Convert.ToDecimal(row["qx_tipo_riesgo_mortalidad_titular"]);
                    worksheet.Cell("BZ" + (1 + i)).Value = Convert.ToDecimal(row["qx_tipo_riesgo_mortalidad_adicional"]);
                    worksheet.Cell("CA" + (1 + i)).Value = Convert.ToDecimal(row["qx_tipo_riesgo_catastrofe_titular"]);
                    worksheet.Cell("CB" + (1 + i)).Value = Convert.ToDecimal(row["qx_tipo_riesgo_catastrofe_adicional"]);

                    worksheet.Cell("CC" + (1 + i)).Value = Convert.ToString(row["t"]);

                    worksheet.Cell("CD" + (1 + i)).Value = Convert.ToDecimal(row["tpx_inicial"]);
                    worksheet.Cell("CE" + (1 + i)).Value = Convert.ToDecimal(row["ajd_qx_monthly"]);
                    worksheet.Cell("CF" + (1 + i)).Value = Convert.ToDecimal(row["lapses"]);
                    worksheet.Cell("CG" + (1 + i)).Value = Convert.ToDecimal(row["invalidez"]);
                    worksheet.Cell("CH" + (1 + i)).Value = Convert.ToDecimal(row["muerte_accidental"]);
                    worksheet.Cell("CI" + (1 + i)).Value = Convert.ToDecimal(row["enfermedades_graves"]);
                    worksheet.Cell("CJ" + (1 + i)).Value = Convert.ToDecimal(row["exoneracion"]);
                    worksheet.Cell("CK" + (1 + i)).Value = Convert.ToDecimal(row["muerte_asegurado_adicional"]);
                    worksheet.Cell("CL" + (1 + i)).Value = Convert.ToDecimal(row["itp_asegurado_adicional"]);
                    worksheet.Cell("CM" + (1 + i)).Value = Convert.ToDecimal(row["tpx_final"]);
                    worksheet.Cell("CN" + (1 + i)).Value = Convert.ToDecimal(row["aux_uno_menos_invalidez"]);
                    worksheet.Cell("CO" + (1 + i)).Value = Convert.ToDecimal(row["aux_uno_menos_enfermedades_graves"]);
                    worksheet.Cell("CP" + (1 + i)).Value = Convert.ToDecimal(row["aux_uno_menos_muerte_aa"]);
                    worksheet.Cell("CQ" + (1 + i)).Value = Convert.ToDecimal(row["aux_uno_menos_itp_aa"]);
                    worksheet.Cell("CR" + (1 + i)).Value = Convert.ToDecimal(row["sx_muerte"]);
                    worksheet.Cell("CS" + (1 + i)).Value = Convert.ToDecimal(row["sx_muerte_accidental"]);
                    worksheet.Cell("CT" + (1 + i)).Value = Convert.ToDecimal(row["sx_invalidez"]);
                    worksheet.Cell("CU" + (1 + i)).Value = Convert.ToDecimal(row["sx_enfermedades_graves"]);
                    worksheet.Cell("CV" + (1 + i)).Value = Convert.ToDecimal(row["sx_exoneracion"]);
                    worksheet.Cell("CW" + (1 + i)).Value = Convert.ToDecimal(row["sx_muerte_aa"]);
                    worksheet.Cell("CX" + (1 + i)).Value = Convert.ToDecimal(row["sx_itp_aa"]);
                    worksheet.Cell("CY" + (1 + i)).Value = Convert.ToDecimal(row["prima_comercial"]);
                    worksheet.Cell("CZ" + (1 + i)).Value = Convert.ToDecimal(row["interes_obtenido"]);
                    worksheet.Cell("DA" + (1 + i)).Value = Convert.ToDecimal(row["frmb_cargo_rescate_total"]);
                    worksheet.Cell("DB" + (1 + i)).Value = Convert.ToDecimal(row["frmb_eg_muerte"]);
                    worksheet.Cell("DC" + (1 + i)).Value = Convert.ToDecimal(row["frmb_eg_itp"]);
                    worksheet.Cell("DD" + (1 + i)).Value = Convert.ToDecimal(row["frmb_eg_ma"]);
                    worksheet.Cell("DE" + (1 + i)).Value = Convert.ToDecimal(row["frmb_eg_exo"]);
                    worksheet.Cell("DF" + (1 + i)).Value = Convert.ToDecimal(row["frmb_eg_enf_graves"]);
                    worksheet.Cell("DG" + (1 + i)).Value = Convert.ToDecimal(row["frmb_eg_muerte_aa"]);
                    worksheet.Cell("DH" + (1 + i)).Value = Convert.ToDecimal(row["frmb_eg_itp_aa"]);
                    worksheet.Cell("DI" + (1 + i)).Value = Convert.ToDecimal(row["frmb_eg_valor_poliza"]);
                    worksheet.Cell("DJ" + (1 + i)).Value = Convert.ToDecimal(row["frmb_eg_comision_basica"]);
                    worksheet.Cell("DK" + (1 + i)).Value = Convert.ToDecimal(row["frmb_eg_comision_excedente"]);
                    worksheet.Cell("DL" + (1 + i)).Value = Convert.ToDecimal(row["frmb_eg_gastos_administrativos"]);
                    worksheet.Cell("DM" + (1 + i)).Value = Convert.ToDecimal(row["frmb_eg_g_gestion_siniestros"]);
                    worksheet.Cell("DN" + (1 + i)).Value = Convert.ToDecimal(row["frmb_eg_gastos_inversiones"]);

                    worksheet.Cell("DO" + (1 + i)).Value = Convert.ToString(row["mv_periodo"]);

                    worksheet.Cell("DP" + (1 + i)).Value = Convert.ToDecimal(row["mv_vtd_vola_mensual"]);
                    worksheet.Cell("DQ" + (1 + i)).Value = Convert.ToDecimal(row["mv_uno_uno_i_t"]);
                    worksheet.Cell("DR" + (1 + i)).Value = Convert.ToDecimal(row["mv_flujo_bop"]);
                    worksheet.Cell("DS" + (1 + i)).Value = Convert.ToDecimal(row["mv_flujo_eop"]);
                    worksheet.Cell("DT" + (1 + i)).Value = Convert.ToDecimal(row["mv_flujo_bel_descontado"]);

                    worksheet.Cell("DU" + (1 + i)).Value = Convert.ToString(row["mtv_periodo"]);

                    worksheet.Cell("DV" + (1 + i)).Value = Convert.ToDecimal(row["mtv_tasa_venta_mensual"]);
                    worksheet.Cell("DW" + (1 + i)).Value = Convert.ToDecimal(row["mtv_uno_uno_i_t"]);
                    worksheet.Cell("DX" + (1 + i)).Value = Convert.ToDecimal(row["mtv_flujo_bop"]);
                    worksheet.Cell("DY" + (1 + i)).Value = Convert.ToDecimal(row["mtv_flujo_eop"]);
                    worksheet.Cell("DZ" + (1 + i)).Value = Convert.ToDecimal(row["mtv_flujo_bel_descontado"]);

                    worksheet.Cell("EA" + (1 + i)).Value = Convert.ToString(row["mtlra_periodo"]);

                    worksheet.Cell("EB" + (1 + i)).Value = Convert.ToDecimal(row["mtlra_tir"]);
                    worksheet.Cell("EC" + (1 + i)).Value = Convert.ToDecimal(row["mtlra_tlra"]);
                    worksheet.Cell("ED" + (1 + i)).Value = Convert.ToDecimal(row["mtlra_flujo_bop"]);
                    worksheet.Cell("ED" + (1 + i)).Value = Convert.ToDecimal(row["mtlra_flujo_eop"]);
                    worksheet.Cell("EE" + (1 + i)).Value = Convert.ToDecimal(row["mtlra_flujo_bel_descontado"]);

                    worksheet.Cell("EF" + (1 + i)).Value = Convert.ToString(row["mtm_periodo"]);

                    worksheet.Cell("EG" + (1 + i)).Value = Convert.ToDecimal(row["mtm_tasa_minima"]);
                    worksheet.Cell("EH" + (1 + i)).Value = Convert.ToDecimal(row["mtm_fd_tasa_minima"]);
                    worksheet.Cell("EI" + (1 + i)).Value = Convert.ToDecimal(row["mtm_flujo_bop"]);
                    worksheet.Cell("EJ" + (1 + i)).Value = Convert.ToDecimal(row["mtm_flujo_eop"]);
                    worksheet.Cell("EK" + (1 + i)).Value = Convert.ToDecimal(row["mtm_flujo_bel_descontado"]);
                    worksheet.Cell("EL" + (1 + i)).Value = Convert.ToDecimal(row["pc_muerte"]);
                    worksheet.Cell("EM" + (1 + i)).Value = Convert.ToDecimal(row["pc_itp"]);
                    worksheet.Cell("EN" + (1 + i)).Value = Convert.ToDecimal(row["pc_ma"]);
                    worksheet.Cell("EO" + (1 + i)).Value = Convert.ToDecimal(row["pc_muerte_aa"]);
                    worksheet.Cell("EP" + (1 + i)).Value = Convert.ToDecimal(row["pc_itp_aa"]);
                    worksheet.Cell("EQ" + (1 + i)).Value = Convert.ToDecimal(row["ci_muerte"]);
                    worksheet.Cell("ER" + (1 + i)).Value = Convert.ToDecimal(row["ci_itp"]);
                    worksheet.Cell("ES" + (1 + i)).Value = Convert.ToDecimal(row["ci_ma"]);
                    worksheet.Cell("ET" + (1 + i)).Value = Convert.ToDecimal(row["ci_muerte_aa"]);
                    worksheet.Cell("EU" + (1 + i)).Value = Convert.ToDecimal(row["ci_itp_aa"]);
                    worksheet.Cell("EV" + (1 + i)).Value = Convert.ToDecimal(row["ce_muerte"]);
                    worksheet.Cell("EW" + (1 + i)).Value = Convert.ToDecimal(row["ce_itp"]);
                    worksheet.Cell("EX" + (1 + i)).Value = Convert.ToDecimal(row["ce_ma"]);
                    worksheet.Cell("EY" + (1 + i)).Value = Convert.ToDecimal(row["ce_muerte_aa"]);
                    worksheet.Cell("EZ" + (1 + i)).Value = Convert.ToDecimal(row["ce_itp_aa"]);
                    worksheet.Cell("FA" + (1 + i)).Value = Convert.ToDecimal(row["ce_menos_ci"]);
                    worksheet.Cell("FB" + (1 + i)).Value = Convert.ToDecimal(row["ce_menos_ci_tasa_minima"]);


                    i = i + 1;
                }

                
                worksheet.ColumnsUsed().AdjustToContents();
                worksheet.Name = "Datos";

                using (MemoryStream stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    stream.Position = 0;
                    ms = stream;
                }
            }

            return ms.ToArray();
        }

    }
}