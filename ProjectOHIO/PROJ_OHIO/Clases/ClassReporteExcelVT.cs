using System;
using System.Data;
using ClosedXML.Excel;
using System.IO;

namespace PROJ_OHIO.Clases
{
    public class ClassReporteExcelVT
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

        public byte[] ExportarResultadoReservaMatematicaVT(DataTable resultado)
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

        public byte[] ExportarDetalleReservaMatematicaVT(DataTable resultado)
        {
            MemoryStream ms = new MemoryStream();

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sample Sheet");

                //// CABECERA
                worksheet.Cell("A1").Value = "Nro. Póliza";
                worksheet.Cell("B1").Value = "Estado de Póliza";
                worksheet.Cell("C1").Value = "fecha_evaluacion";

                worksheet.Cell("D1").Value = "t_poliza";
                worksheet.Cell("E1").Value = "mes_poliza";
                worksheet.Cell("F1").Value = "mes_pago";
                worksheet.Cell("G1").Value = "edad_nearest_to_the_birthday";
                worksheet.Cell("H1").Value = "edad_last_birthday";
                worksheet.Cell("I1").Value = "mejora_mortalidad";

                worksheet.Cell("J1").Value = "prima_comercial";
                worksheet.Cell("K1").Value = "qx_anual";
                worksheet.Cell("L1").Value = "qx_base";
                worksheet.Cell("M1").Value = "ajdqxmonthly";
                worksheet.Cell("N1").Value = "morb_qx_itp";

                worksheet.Cell("O1").Value = "morb_qx_base_itp";
                worksheet.Cell("P1").Value = "ajd_morb_qx_monthly_itp";
                worksheet.Cell("Q1").Value = "morb_qx_ma";
                worksheet.Cell("R1").Value = "morb_qx_base_ma";
                worksheet.Cell("S1").Value = "ajd_morb_qx_monthly_ma";
                worksheet.Cell("T1").Value = "morb_qx_enf_grave";
                worksheet.Cell("U1").Value = "morb_qx_base_enf_grave";
                worksheet.Cell("V1").Value = "ajd_morb_qx_monthly_enf_grave";
                worksheet.Cell("W1").Value = "morb_qx_exoneracion";
                worksheet.Cell("X1").Value = "morb_qx_base_exoneracion";
                worksheet.Cell("Y1").Value = "ajd_morb_qx_monthly_exoneracion";
                worksheet.Cell("Z1").Value = "t";

                worksheet.Cell("AA1").Value = "tpx_inicial";
                worksheet.Cell("AB1").Value = "ajd_qx_monthly";
                worksheet.Cell("AC1").Value = "lapses";
                worksheet.Cell("AD1").Value = "invalidez";
                worksheet.Cell("AE1").Value = "muerte_accidental";
                worksheet.Cell("AF1").Value = "enfermedades_graves";
                worksheet.Cell("AG1").Value = "exoneracion";
                worksheet.Cell("AH1").Value = "tpx_final";
                worksheet.Cell("AI1").Value = "aux_uno_menos_invalidez";
                worksheet.Cell("AJ1").Value = "aux_uno_menos_enfermedades_graves";
                worksheet.Cell("AK1").Value = "edad_meses_titular_mortalidad";
                worksheet.Cell("AL1").Value = "edad_anios_t_titular_mortalidad";

                worksheet.Cell("AM1").Value = "edad_anios_t_mas_1_titular_mortalidad";
                worksheet.Cell("AN1").Value = "lx_t_titular_mortalidad";
                worksheet.Cell("AO1").Value = "lx_t_mas_1_titular_mortalidad";
                worksheet.Cell("AP1").Value = "lx_probabilidad_titular_mortalidad";
                worksheet.Cell("AQ1").Value = "qx_mortalidad";
                worksheet.Cell("AR1").Value = "edad_meses_titular_catastrofe";
                worksheet.Cell("AS1").Value = "edad_anios_t_titular_catastrofe";
                worksheet.Cell("AT1").Value = "edad_anios_t_mas_1_titular_catastrofe";
                worksheet.Cell("AU1").Value = "lx_t_titular_catastrofe";
                worksheet.Cell("AV1").Value = "lx_t_mas_1_titular_catastrofe";
                worksheet.Cell("AW1").Value = "lx_probabilidad_titular_catastrofe";
                worksheet.Cell("AX1").Value = "qx_catastrofe";

                worksheet.Cell("AY1").Value = "pc_muerte";
                worksheet.Cell("AZ1").Value = "pc_itp";
                worksheet.Cell("BA1").Value = "pc_ma";
                worksheet.Cell("BB1").Value = "pc_enf_grave";
                worksheet.Cell("BC1").Value = "pc_exo";
                worksheet.Cell("BD1").Value = "suma_asegurada_muerte";
                worksheet.Cell("BE1").Value = "suma_asegurada_itp";
                worksheet.Cell("BF1").Value = "suma_asegurada_ma";
                worksheet.Cell("BG1").Value = "frmb_ing_muerte";
                worksheet.Cell("BH1").Value = "frmb_ing_itp";
                worksheet.Cell("BI1").Value = "frmb_ing_ma";
                worksheet.Cell("BJ1").Value = "frmb_ing_enf_graves";

                worksheet.Cell("BK1").Value = "frmb_ing_exo";
                worksheet.Cell("BL1").Value = "frmb_eg_muerte";
                worksheet.Cell("BM1").Value = "frmb_eg_itp";
                worksheet.Cell("BN1").Value = "frmb_eg_ma";
                worksheet.Cell("BO1").Value = "frmb_eg_enf_graves";
                worksheet.Cell("BP1").Value = "frmb_eg_exo";
                worksheet.Cell("BQ1").Value = "frmb_eg_rescate";
                worksheet.Cell("BR1").Value = "frmb_eg_gastos_adquisicion";
                worksheet.Cell("BS1").Value = "frmb_eg_gastos_administrativos";
                worksheet.Cell("BT1").Value = "frmb_eg_g_gestion_siniestros";
                worksheet.Cell("BU1").Value = "frmb_eg_gastos_inversiones";

                worksheet.Cell("BV1").Value = "mv_periodo";
                worksheet.Cell("BW1").Value = "mv_vtd_vola_mensual";
                worksheet.Cell("BX1").Value = "mv_uno_uno_i_t";
                worksheet.Cell("BY1").Value = "mv_flujo_bop";
                worksheet.Cell("BZ1").Value = "mv_flujo_eop";
                worksheet.Cell("CA1").Value = "mv_flujo_bel_descontado";

                worksheet.Cell("CB1").Value = "mtv_periodo";
                worksheet.Cell("CC1").Value = "mtv_tasa_venta_mensual";
                worksheet.Cell("CD1").Value = "mtv_uno_uno_i_t";
                worksheet.Cell("CE1").Value = "mtv_flujo_bop";
                worksheet.Cell("CF1").Value = "mtv_flujo_eop";
                worksheet.Cell("CG1").Value = "mtv_flujo_bel_descontado";

                worksheet.Cell("CH1").Value = "mtlra_periodo";
                worksheet.Cell("CI1").Value = "mtlra_tir";
                worksheet.Cell("CJ1").Value = "mtlra_tlra";
                worksheet.Cell("CK1").Value = "mtlra_flujo_bop";
                worksheet.Cell("CL1").Value = "mtlra_flujo_eop";
                worksheet.Cell("CM1").Value = "mtlra_flujo_bel_descontado";

                worksheet.Cell("CN1").Value = "mtm_periodo";
                worksheet.Cell("CO1").Value = "mtm_tasa_minima";
                worksheet.Cell("CP1").Value = "mtm_fd_tasa_minima";
                worksheet.Cell("CQ1").Value = "mtm_flujo_bop";
                worksheet.Cell("CR1").Value = "mtm_flujo_eop";
                worksheet.Cell("CS1").Value = "mtm_flujo_bel_descontado";
                worksheet.Cell("CT1").Value = "pce_muerte";

                worksheet.Cell("CU1").Value = "pce_itp";
                worksheet.Cell("CV1").Value = "pce_ma";
                worksheet.Cell("CW1").Value = "ci_muerte";
                worksheet.Cell("CX1").Value = "ci_itp";
                worksheet.Cell("CY1").Value = "ci_ma";
                worksheet.Cell("CZ1").Value = "ce_muerte";
                worksheet.Cell("DA1").Value = "ce_itp";
                worksheet.Cell("DB1").Value = "ce_ma";
                worksheet.Cell("DC1").Value = "ce_menos_ci";
                worksheet.Cell("DD1").Value = "ce_menos_ci_tasa_minima";

                
                ////worksheet.Cell("A1:J1").Style.Font.FontColor = XLColor.Red;
                worksheet.Range("A1:DD1").Style.Fill.BackgroundColor = XLColor.Gray;
                worksheet.Range("A1:DD1").Style.Font.Bold = true;
                worksheet.Range("A1:DD1").Style.Font.FontColor = XLColor.White;


                int i = 1;
                foreach (DataRow row in resultado.Rows)
                {

                    worksheet.Cell("A" + (1 + i)).Value = Convert.ToString(row["NumeroPoliza"]);
                    worksheet.Cell("B" + (1 + i)).Value = Convert.ToString(row["EstadoPoliza"]);
                    worksheet.Cell("C" + (1 + i)).Value = Convert.ToString(row["fecha_evaluacion"]);

                    worksheet.Cell("D" + (1 + i)).Value = Convert.ToString(row["t_poliza"]);
                    worksheet.Cell("E" + (1 + i)).Value = Convert.ToString(row["mes_poliza"]);
                    worksheet.Cell("F" + (1 + i)).Value = Convert.ToString(row["mes_pago"]);
                    worksheet.Cell("G" + (1 + i)).Value = Convert.ToString(row["edad_nearest_to_the_birthday"]);
                    worksheet.Cell("H" + (1 + i)).Value = Convert.ToString(row["edad_last_birthday"]);
                    worksheet.Cell("I" + (1 + i)).Value = Convert.ToString(row["mejora_mortalidad"]);


                    worksheet.Cell("J" + (1 + i)).Value = Convert.ToDecimal(row["prima_comercial"]);
                    worksheet.Cell("K" + (1 + i)).Value = Convert.ToDecimal(row["qx_anual"]);
                    worksheet.Cell("L" + (1 + i)).Value = Convert.ToDecimal(row["qx_base"]);
                    worksheet.Cell("M" + (1 + i)).Value = Convert.ToDecimal(row["ajdqxmonthly"]);
                    worksheet.Cell("N" + (1 + i)).Value = Convert.ToDecimal(row["morb_qx_itp"]);
                    worksheet.Cell("O" + (1 + i)).Value = Convert.ToDecimal(row["morb_qx_base_itp"]);
                    worksheet.Cell("P" + (1 + i)).Value = Convert.ToDecimal(row["ajd_morb_qx_monthly_itp"]);
                    worksheet.Cell("Q" + (1 + i)).Value = Convert.ToDecimal(row["morb_qx_ma"]);
                    worksheet.Cell("R" + (1 + i)).Value = Convert.ToDecimal(row["morb_qx_base_ma"]);
                    worksheet.Cell("S" + (1 + i)).Value = Convert.ToDecimal(row["ajd_morb_qx_monthly_ma"]);
                    worksheet.Cell("T" + (1 + i)).Value = Convert.ToDecimal(row["morb_qx_enf_grave"]);
                    worksheet.Cell("U" + (1 + i)).Value = Convert.ToDecimal(row["morb_qx_base_enf_grave"]);
                    worksheet.Cell("V" + (1 + i)).Value = Convert.ToDecimal(row["ajd_morb_qx_monthly_enf_grave"]);
                    worksheet.Cell("W" + (1 + i)).Value = Convert.ToDecimal(row["morb_qx_exoneracion"]);
                    worksheet.Cell("X" + (1 + i)).Value = Convert.ToDecimal(row["morb_qx_base_exoneracion"]);
                    worksheet.Cell("Y" + (1 + i)).Value = Convert.ToDecimal(row["ajd_morb_qx_monthly_exoneracion"]);
                    worksheet.Cell("Z" + (1 + i)).Value = Convert.ToDecimal(row["t"]);
                    worksheet.Cell("AA" + (1 + i)).Value = Convert.ToDecimal(row["tpx_inicial"]);
                    worksheet.Cell("AB" + (1 + i)).Value = Convert.ToDecimal(row["ajd_qx_monthly"]);
                    worksheet.Cell("AC" + (1 + i)).Value = Convert.ToDecimal(row["lapses"]);
                    worksheet.Cell("AD" + (1 + i)).Value = Convert.ToDecimal(row["invalidez"]);
                    worksheet.Cell("AE" + (1 + i)).Value = Convert.ToDecimal(row["muerte_accidental"]);
                    worksheet.Cell("AF" + (1 + i)).Value = Convert.ToDecimal(row["enfermedades_graves"]);
                    worksheet.Cell("AG" + (1 + i)).Value = Convert.ToDecimal(row["exoneracion"]);
                    worksheet.Cell("AH" + (1 + i)).Value = Convert.ToDecimal(row["tpx_final"]);
                    worksheet.Cell("AI" + (1 + i)).Value = Convert.ToDecimal(row["aux_uno_menos_invalidez"]);
                    worksheet.Cell("AJ" + (1 + i)).Value = Convert.ToDecimal(row["aux_uno_menos_enfermedades_graves"]);
                    worksheet.Cell("AK" + (1 + i)).Value = Convert.ToDecimal(row["edad_meses_titular_mortalidad"]);
                    worksheet.Cell("AL" + (1 + i)).Value = Convert.ToDecimal(row["edad_anios_t_titular_mortalidad"]);
                    worksheet.Cell("AM" + (1 + i)).Value = Convert.ToDecimal(row["edad_anios_t_mas_1_titular_mortalidad"]);
                    worksheet.Cell("AN" + (1 + i)).Value = Convert.ToDecimal(row["lx_t_titular_mortalidad"]);
                    worksheet.Cell("AO" + (1 + i)).Value = Convert.ToDecimal(row["lx_t_mas_1_titular_mortalidad"]);
                    worksheet.Cell("AP" + (1 + i)).Value = Convert.ToDecimal(row["lx_probabilidad_titular_mortalidad"]);
                    worksheet.Cell("AQ" + (1 + i)).Value = Convert.ToDecimal(row["qx_mortalidad"]);

                    worksheet.Cell("AR" + (1 + i)).Value = Convert.ToString(row["edad_meses_titular_catastrofe"]);
                    worksheet.Cell("AS" + (1 + i)).Value = Convert.ToString(row["edad_anios_t_titular_catastrofe"]);
                    worksheet.Cell("AT" + (1 + i)).Value = Convert.ToString(row["edad_anios_t_mas_1_titular_catastrofe"]);

                    worksheet.Cell("AU" + (1 + i)).Value = Convert.ToDecimal(row["lx_t_titular_catastrofe"]);
                    worksheet.Cell("AV" + (1 + i)).Value = Convert.ToDecimal(row["lx_t_mas_1_titular_catastrofe"]);
                    worksheet.Cell("AW" + (1 + i)).Value = Convert.ToDecimal(row["lx_probabilidad_titular_catastrofe"]);
                    worksheet.Cell("AX" + (1 + i)).Value = Convert.ToDecimal(row["qx_catastrofe"]);
                    worksheet.Cell("AY" + (1 + i)).Value = Convert.ToDecimal(row["pc_muerte"]);

                    worksheet.Cell("AZ" + (1 + i)).Value = Convert.ToDecimal(row["pc_itp"]);
                    worksheet.Cell("BA" + (1 + i)).Value = Convert.ToDecimal(row["pc_ma"]);
                    worksheet.Cell("BB" + (1 + i)).Value = Convert.ToDecimal(row["pc_enf_grave"]);
                    worksheet.Cell("BC" + (1 + i)).Value = Convert.ToDecimal(row["pc_exo"]);
                    worksheet.Cell("BD" + (1 + i)).Value = Convert.ToDecimal(row["suma_asegurada_muerte"]);
                    worksheet.Cell("BE" + (1 + i)).Value = Convert.ToDecimal(row["suma_asegurada_itp"]);
                    worksheet.Cell("BF" + (1 + i)).Value = Convert.ToDecimal(row["suma_asegurada_ma"]);
                    worksheet.Cell("BG" + (1 + i)).Value = Convert.ToDecimal(row["frmb_ing_muerte"]);
                    worksheet.Cell("BH" + (1 + i)).Value = Convert.ToDecimal(row["frmb_ing_itp"]);


                    worksheet.Cell("BI" + (1 + i)).Value = Convert.ToDecimal(row["frmb_ing_ma"]);
                    worksheet.Cell("BJ" + (1 + i)).Value = Convert.ToDecimal(row["frmb_ing_enf_graves"]);
                    worksheet.Cell("BK" + (1 + i)).Value = Convert.ToDecimal(row["frmb_ing_exo"]);

                    worksheet.Cell("BL" + (1 + i)).Value = Convert.ToDecimal(row["frmb_eg_muerte"]);
                    worksheet.Cell("BM" + (1 + i)).Value = Convert.ToDecimal(row["frmb_eg_itp"]);
                    worksheet.Cell("BN" + (1 + i)).Value = Convert.ToDecimal(row["frmb_eg_ma"]);

                    worksheet.Cell("BO" + (1 + i)).Value = Convert.ToDecimal(row["frmb_eg_enf_graves"]);
                    worksheet.Cell("BP" + (1 + i)).Value = Convert.ToDecimal(row["frmb_eg_exo"]);
                    worksheet.Cell("BQ" + (1 + i)).Value = Convert.ToDecimal(row["frmb_eg_rescate"]);


                    worksheet.Cell("BR" + (1 + i)).Value = Convert.ToDecimal(row["frmb_eg_gastos_adquisicion"]);
                    worksheet.Cell("BS" + (1 + i)).Value = Convert.ToDecimal(row["frmb_eg_gastos_administrativos"]);
                    worksheet.Cell("BT" + (1 + i)).Value = Convert.ToDecimal(row["frmb_eg_g_gestion_siniestros"]);
                    worksheet.Cell("BU" + (1 + i)).Value = Convert.ToDecimal(row["frmb_eg_gastos_inversiones"]);
                    
                    worksheet.Cell("BV" + (1 + i)).Value = Convert.ToString(row["mv_periodo"]);

                    worksheet.Cell("BW" + (1 + i)).Value = Convert.ToDecimal(row["mv_vtd_vola_mensual"]);
                    worksheet.Cell("BX" + (1 + i)).Value = Convert.ToDecimal(row["mv_uno_uno_i_t"]); //// error null
                    worksheet.Cell("BY" + (1 + i)).Value = Convert.ToDecimal(row["mv_flujo_bop"]);
                    worksheet.Cell("BZ" + (1 + i)).Value = Convert.ToDecimal(row["mv_flujo_eop"]);
                    worksheet.Cell("CA" + (1 + i)).Value = Convert.ToDecimal(row["mv_flujo_bel_descontado"]);

                    worksheet.Cell("CB" + (1 + i)).Value = Convert.ToString(row["mtv_periodo"]);

                    worksheet.Cell("CC" + (1 + i)).Value = Convert.ToDecimal(row["mtv_tasa_venta_mensual"]);
                    worksheet.Cell("CD" + (1 + i)).Value = Convert.ToDecimal(row["mtv_uno_uno_i_t"]);
                    worksheet.Cell("CE" + (1 + i)).Value = Convert.ToDecimal(row["mtv_flujo_bop"]);
                    worksheet.Cell("CF" + (1 + i)).Value = Convert.ToDecimal(row["mtv_flujo_eop"]);
                    worksheet.Cell("CG" + (1 + i)).Value = Convert.ToDecimal(row["mtv_flujo_bel_descontado"]);

                    worksheet.Cell("CH" + (1 + i)).Value = Convert.ToString(row["mtlra_periodo"]);

                    worksheet.Cell("CI" + (1 + i)).Value = Convert.ToDecimal(row["mtlra_tir"]);
                    worksheet.Cell("CJ" + (1 + i)).Value = Convert.ToDecimal(row["mtlra_tlra"]);
                    worksheet.Cell("CK" + (1 + i)).Value = Convert.ToDecimal(row["mtlra_flujo_bop"]);
                    worksheet.Cell("CL" + (1 + i)).Value = Convert.ToDecimal(row["mtlra_flujo_eop"]);
                    worksheet.Cell("CM" + (1 + i)).Value = Convert.ToDecimal(row["mtlra_flujo_bel_descontado"]);

                    worksheet.Cell("CN" + (1 + i)).Value = Convert.ToString(row["mtm_periodo"]);

                    worksheet.Cell("CO" + (1 + i)).Value = Convert.ToDecimal(row["mtm_tasa_minima"]);
                    worksheet.Cell("CP" + (1 + i)).Value = Convert.ToDecimal(row["mtm_fd_tasa_minima"]);
                    worksheet.Cell("CQ" + (1 + i)).Value = Convert.ToDecimal(row["mtm_flujo_bop"]);
                    worksheet.Cell("CR" + (1 + i)).Value = Convert.ToDecimal(row["mtm_flujo_eop"]);
                    worksheet.Cell("CS" + (1 + i)).Value = Convert.ToDecimal(row["mtm_flujo_bel_descontado"]);
                    worksheet.Cell("CT" + (1 + i)).Value = Convert.ToDecimal(row["pce_muerte"]);
                    worksheet.Cell("CU" + (1 + i)).Value = Convert.ToDecimal(row["pce_itp"]);
                    worksheet.Cell("CV" + (1 + i)).Value = Convert.ToDecimal(row["pce_ma"]);
                    worksheet.Cell("CW" + (1 + i)).Value = Convert.ToDecimal(row["ci_muerte"]);
                    worksheet.Cell("CX" + (1 + i)).Value = Convert.ToDecimal(row["ci_itp"]);
                    worksheet.Cell("CY" + (1 + i)).Value = Convert.ToDecimal(row["ci_ma"]);
                    worksheet.Cell("CZ" + (1 + i)).Value = Convert.ToDecimal(row["ce_muerte"]);
                    worksheet.Cell("DA" + (1 + i)).Value = Convert.ToDecimal(row["ce_itp"]);
                    worksheet.Cell("DB" + (1 + i)).Value = Convert.ToDecimal(row["ce_ma"]);
                    worksheet.Cell("DC" + (1 + i)).Value = Convert.ToDecimal(row["ce_menos_ci"]);
                    worksheet.Cell("DD" + (1 + i)).Value = Convert.ToDecimal(row["ce_menos_ci_tasa_minima"]);

                    

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