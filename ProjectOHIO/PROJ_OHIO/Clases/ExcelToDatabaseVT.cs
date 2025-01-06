using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Globalization;
using System.IO;
using ExcelDataReader;

namespace PROJ_OHIO.Clases
{
    public class ExcelToDatabaseVT
    {
		public bool PolizasDatabase(DataTable dt, string moneda, int anio, int mes)
		{
			bool result = false;
			string mensaje = "";

			ModeloNegocio.MyNegocio nObj = new ModeloNegocio.MyNegocio();

			int anio_periodo;
			int mes_periodo;
			string Moneda;
			string NumeroPoliza;
			int NumeroSolicitud;
			int EstadoPoliza;
			DateTime? FechaInicioVigente;
			DateTime? FechaCancelacionPoliza;
			string MotivoCancelacion;
			DateTime? FechaUltimoEndoso;
			DateTime? FechaSuspensionPoliza;
			string TipoSeguro;
			string PeriodicidadPrima;
			string NombreContratante;
			int TipoDocumentoContratante;
			string NumeroDocumentoContratante;
			string ContratantePersonaJuridica;
			int Temporalidad;
			int Duracion;
			int AniosPago;
			string MetodoPagoPoliza;
			decimal PrimaBasica1erAnio;
			decimal PrimaBasicaDemasAnios;
			decimal PrimaMinima;
			decimal PrimaMaxima;
			decimal PrimaPactadaAnual;
			string NombreAseguradoTitular;
			int TipoDocumentoAseguradoTitular;
			string NumeroDocumentoAseguradoTitular;
			DateTime? FechaNacimientoAseguradoTitular;
			int EdadActuarialAseguradoTitular;
			string SexoAseguradoTitular;
			string ClaseRiesgoAseguradoTitular;
			string FumadorAseguradoTitular;
			decimal SumaAseguradaActualCoberturaMuerte;
			decimal SobremortalidadCoberturaMuerte;
			decimal SobreprimaCoberturaMuerte;
			decimal SumaAseguradaActualCoberturaITP;
			decimal SobremortalidadCoberturaITP;
			decimal SobreprimaCoberturaITP;
			decimal SumaAseguradaActualCoberturaMuerteAccidental;
			decimal SobremortalidadCoberturaMuerteAccidental;
			decimal SobreprimaCoberturaMuerteAccidental;
			decimal SumaAseguradaActualCoberturaEnfermedadesGraves;
			string CubiertoCoberturaExoneracion;
			string NombreAseguradoAdicional;
			int TipoDocumentoAseguradoAdicional;
			string NumeroDocumentoAseguradoAdicional;
			DateTime? FechaNacimientoAseguradoAdicional;
			int EdadActuarialEmisionAseguradoAdicional;
			string SexoAseguradoAdicional;
			string ClaseRiesgoAseguradoAdicional;
			string FumadorAseguradoAdicional;
			string RelacionConTitular;
			decimal SumaAseguradaActualCoberturaMuerteAseguradoAdicional;
			decimal SobremortalidadCoberturaMuerteAseguradoAdicional;
			decimal SobreprimaCoberturaMuerteAseguradoAdicional;
			decimal SumaAseguradaActualCoberturaITPAseguradoAdicional;
			decimal SobremortalidadCoberturaITPAseguradoAdicional;
			decimal SobreprimaCoberturaITPAseguradoAdicional;
			decimal TasaInteresAcreditadoMes;
			string CodigoCorredor;
			string NombreCorredor;
			int TipoIntermediario;
			string IgvComision;
			int AnioPolizaNumero;
			decimal PorcentajeComisionBrokerPBAnioActual;
			decimal PorcentajeComisionBrokerPEAnioActual;
			decimal ComisionSinIGVPagadoDuranteMes;
			decimal IgvComision2;
			decimal ComisionDevengadaDAC;
			decimal GastoMedicoDevengado;

			decimal ValorPolizaBOP;
			decimal PrimaPagada;
			decimal Interes;
			decimal RescatesParciales;
			decimal CargoPorRescateParcial;
			decimal GastoAdministrativoPB;
			decimal GastoAdministrativoPexc;
			decimal GastoAdministrativoDisminucionSA;
			decimal COI_Muerte;
			decimal COI_ITP;
			decimal COI_MuerteAccidental;
			decimal COI_EnfermedadesGraves;
			decimal COI_Exoneracion;
			decimal COI_MuerteAA;
			decimal COI_ITP_AA;
			decimal CargoPorRescateTotalCancelacion;
			decimal COIS_PendientesPago;
			decimal ValorPolizaEOP;
			decimal ValorRescate;

			try
			{
				int posicion = 0;
				string fecha_null = "01/01/1000";

				//Pass the filepath and filename to the StreamWriter Constructor
				//StreamWriter sw = new StreamWriter("C:\\Backup\\Test.txt");

				foreach (DataRow row in dt.Rows)
				{
					//tblObj.Name = row["Name"].ToString();
					//tblObj.Name = row["Address"].ToString();
					//tblObj.Salary = (int)row["Salary"];
					//tblObj.Age = (int)row["Age"];
					posicion += 1;

					if (posicion >= 6 /*&& posicion<=100*/)
					{
						anio_periodo = anio;
						mes_periodo = mes;
						Moneda = moneda;
						NumeroPoliza = Convert.ToString(row[2]);
						NumeroSolicitud = Convert.ToInt32(row[1]);

						int _estado = -1;
						switch (row[3].ToString().ToUpper())
						{
							case "ANULADA":
								_estado = 0;
								break;
							case "CANCELACIÓN POR FALTA DE PAGO":
								_estado = 1;
								break;
							case "CANCELACION POR FALTA DE PAGO":
								_estado = 1;
								break;
							case "CANCELACIÓN POR RESCATE":
								_estado = 2;
								break;
							case "CANCELACION POR RESCATE":
								_estado = 2;
								break;
							case "CANCELADO POR RESCATE":
								_estado = 3;
								break;
							case "COBERTURA SUSPENDIDA":
								_estado = 4;
								break;
							case "PERIODO DE GRACIA":
								_estado = 5;
								break;
							case "PÓLIZA VIGENTE":
								_estado = 6;
								break;
							case "POLIZA VIGENTE":
								_estado = 6;
								break;
							case "SOLICITUD APROBADA EN COBRANZA":
								_estado = 7;
								break;
						}

						//switch (row[3].ToString().ToUpper())
						//{
						//	case "Anulada":
						//		_estado = 0;
						//		break;
						//	case "Cancelación por falta de pago":
						//		_estado = 1;
						//		break;
						//	case "Cancelacion por falta de pago":
						//		_estado = 1;
						//		break;
						//	case "Cancelación por rescate":
						//		_estado = 2;
						//		break;
						//	case "Cancelacion por rescate":
						//		_estado = 2;
						//		break;
						//	case "Cancelado por rescate":
						//		_estado = 3;
						//		break;
						//	case "Cobertura suspendida":
						//		_estado = 4;
						//		break;
						//	case "Periodo de gracia":
						//		_estado = 5;
						//		break;
						//	case "Póliza Vigente":
						//		_estado = 6;
						//		break;
						//	case "Poliza Vigente":
						//		_estado = 6;
						//		break;
						//	case "Solicitud aprobada en cobranza":
						//		_estado = 7;
						//		break;
						//}

						EstadoPoliza = _estado;
						//sw.WriteLine("Data FechaInicioVigente :: " + NumeroPoliza + "  ::  " + row[4].ToString());
						//sw.WriteLine("-------------------------------------------");
						FechaInicioVigente = (row[4].ToString() == "" ? Convert.ToDateTime(fecha_null) : Convert.ToDateTime(row[4]));



						//sw.WriteLine("Data FechaCancelacionPoliza :: " + NumeroPoliza + "  ::  " + row[5].ToString());
						//sw.WriteLine("-------------------------------------------");
						FechaCancelacionPoliza = (row[5].ToString() == "" ? Convert.ToDateTime(fecha_null) : Convert.ToDateTime(row[5]));

						MotivoCancelacion = row[6].ToString();

						//sw.WriteLine("Data FechaUltimoEndoso :: " + NumeroPoliza + "  ::  " + row[7].ToString());
						//sw.WriteLine("-------------------------------------------");
						FechaUltimoEndoso = (row[7].ToString() == "" ? Convert.ToDateTime(fecha_null) : Convert.ToDateTime(row[7]));

						//sw.WriteLine("Data FechaSuspensionPoliza :: " + NumeroPoliza + "  ::  " + row[8].ToString());
						//sw.WriteLine("-------------------------------------------");
						FechaSuspensionPoliza = (row[8].ToString() == "" ? Convert.ToDateTime(fecha_null) : Convert.ToDateTime(row[8]));
						TipoSeguro = row[9].ToString();

						string _periodicidad = "";
						switch (row[10].ToString())
						{
							case "Anual":
								_periodicidad = "A";
								break;
							case "Mensual":
								_periodicidad = "M";
								break;
							case "Semestral":
								_periodicidad = "S";
								break;
							case "Trimestral":
								_periodicidad = "T";
								break;
						}
						PeriodicidadPrima = _periodicidad;

						NombreContratante = row[11].ToString();

						int _tipoDocCont = -1;
						switch (row[12].ToString())
						{
							case "DNI":
								_tipoDocCont = 0;
								break;
							case "RUC":
								_tipoDocCont = 1;
								break;
							case "Carnet de extranjeria":
								_tipoDocCont = 2;
								break;
							case "Carnét de extranjería":
								_tipoDocCont = 2;
								break;
						}
						TipoDocumentoContratante = _tipoDocCont;

						NumeroDocumentoContratante = row[13].ToString();
						ContratantePersonaJuridica = row[14].ToString().Substring(0, 1);
						Temporalidad = Convert.ToInt32(row[15]);
						Duracion = Convert.ToInt32(row[16]);
						AniosPago = Convert.ToInt32(row[17]);
						MetodoPagoPoliza = row[18].ToString().Substring(0, 1);

						//Console.WriteLine("mostrando data :: " + NumeroPoliza );
						//Console.WriteLine(row[19].ToString());
						//Console.ReadKey();
						//Console.WriteLine("mostrando data :: " + NumeroPoliza);
						//this.Response.Write(" ERROR :: " + ex.Message);
						//ViewBag.software = software;

						//Console.Write("aaaaaaaaa");
						//System.Diagnostics.Debug.WriteLine("Data :: " + NumeroPoliza + "  ::  " + row[19].ToString());
						//System.Diagnostics.Debug.Print("ada");


						//Write a line of text
						////sw.WriteLine("Data PrimaBasica1erAnio :: " + NumeroPoliza + "  ::  " + row[19].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""));						
						//sw.WriteLine("-------------------------------------------");						
						PrimaBasica1erAnio = Convert.ToDecimal(row[19].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						////sw.WriteLine("PrimaBasica1erAnio :: ::  " + PrimaBasica1erAnio.ToString());
						//sw.WriteLine("************************************************");



						//sw.WriteLine("Data PrimaBasicaDemasAnios :: " + NumeroPoliza + "  ::  " + row[20].ToString());
						//sw.WriteLine("-------------------------------------------");
						PrimaBasicaDemasAnios = Convert.ToDecimal(row[20].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						//sw.WriteLine("PrimaBasicaDemasAnios :: ::  " + PrimaBasicaDemasAnios.ToString());
						//sw.WriteLine("************************************************");


						//sw.WriteLine("Data PrimaMinima :: " + NumeroPoliza + "  ::  " + row[21].ToString());
						//sw.WriteLine("-------------------------------------------");

						PrimaMinima = Convert.ToDecimal(row[21].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));

						//sw.WriteLine("PrimaMinima :: ::  " + PrimaMinima.ToString());
						//sw.WriteLine("************************************************");



						//sw.WriteLine("Data PrimaMaxima :: " + NumeroPoliza + "  ::  " + row[22].ToString());
						//sw.WriteLine("-------------------------------------------");

						PrimaMaxima = Convert.ToDecimal(row[22].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						//sw.WriteLine("PrimaMaxima :: ::  " + PrimaMaxima.ToString());
						//sw.WriteLine("************************************************");


						//sw.WriteLine("Data PrimaPactadaAnual :: " + NumeroPoliza + "  ::  " + row[23].ToString());
						//sw.WriteLine("-------------------------------------------");
						PrimaPactadaAnual = Convert.ToDecimal(row[23].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						//sw.WriteLine("PrimaMaxima :: ::  " + PrimaMaxima.ToString());
						//sw.WriteLine("************************************************");


						//sw.WriteLine("Data NombreAseguradoTitular :: " + NumeroPoliza + "  ::  " + row[24].ToString());
						//sw.WriteLine("-------------------------------------------");

						NombreAseguradoTitular = row[24].ToString();

						int _tipoDocCont2 = -1;
						switch (row[25].ToString())
						{
							case "DNI":
								_tipoDocCont2 = 0;
								break;
							case "RUC":
								_tipoDocCont2 = 1;
								break;
							case "Carnet de extranjeria":
								_tipoDocCont2 = 2;
								break;
							case "Carnét de extranjería":
								_tipoDocCont2 = 2;
								break;
						}
						TipoDocumentoAseguradoTitular = _tipoDocCont2;
						NumeroDocumentoAseguradoTitular = row[26].ToString();
						FechaNacimientoAseguradoTitular = (row[27].ToString() == "" ? Convert.ToDateTime(fecha_null) : Convert.ToDateTime(row[27]));
						EdadActuarialAseguradoTitular = Convert.ToInt32(row[28]);
						SexoAseguradoTitular = row[29].ToString().Substring(0, 1);
						ClaseRiesgoAseguradoTitular = row[30].ToString().Substring(0, 1);
						FumadorAseguradoTitular = row[31].ToString().Substring(0, 1);

						SumaAseguradaActualCoberturaMuerte = Convert.ToDecimal(row[32].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						SobremortalidadCoberturaMuerte = Convert.ToDecimal(row[33].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						SobreprimaCoberturaMuerte = Convert.ToDecimal(row[34].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						SumaAseguradaActualCoberturaITP = Convert.ToDecimal(row[35].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						SobremortalidadCoberturaITP = Convert.ToDecimal(row[36].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						SobreprimaCoberturaITP = Convert.ToDecimal(row[37].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						SumaAseguradaActualCoberturaMuerteAccidental = Convert.ToDecimal(row[38].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						SobremortalidadCoberturaMuerteAccidental = Convert.ToDecimal(row[39].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						SobreprimaCoberturaMuerteAccidental = Convert.ToDecimal(row[40].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						SumaAseguradaActualCoberturaEnfermedadesGraves = Convert.ToDecimal(row[41].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						CubiertoCoberturaExoneracion = row[42].ToString().Substring(0, 1);
						NombreAseguradoAdicional = row[43].ToString();


						int _tipoDocCont3 = -1;
						switch (row[44].ToString())
						{
							case "DNI":
								_tipoDocCont3 = 0;
								break;
							case "RUC":
								_tipoDocCont3 = 1;
								break;
							case "Carnet de extranjeria":
								_tipoDocCont3 = 2;
								break;
							case "Carnét de extranjería":
								_tipoDocCont3 = 2;
								break;
						}
						TipoDocumentoAseguradoAdicional = _tipoDocCont3;
						NumeroDocumentoAseguradoAdicional = row[45].ToString();
						FechaNacimientoAseguradoAdicional = (row[46].ToString() == "" ? Convert.ToDateTime(fecha_null) : Convert.ToDateTime(row[46]));
						EdadActuarialEmisionAseguradoAdicional = (row[47].ToString() == "" ? 0 : Convert.ToInt32(row[47]));  //Convert.ToInt32(row[47]);
						SexoAseguradoAdicional = (row[48].ToString() == "" ? "" : row[48].ToString().Substring(0, 1));  //row[48].ToString().Substring(0, 1);
						ClaseRiesgoAseguradoAdicional = (row[49].ToString() == "" ? "" : row[49].ToString().Substring(0, 1));
						FumadorAseguradoAdicional = (row[50].ToString() == "" ? "" : row[50].ToString().Substring(0, 1));
						RelacionConTitular = row[51].ToString();
						SumaAseguradaActualCoberturaMuerteAseguradoAdicional = Convert.ToDecimal(row[52].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						SobremortalidadCoberturaMuerteAseguradoAdicional = Convert.ToDecimal(row[53].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						SobreprimaCoberturaMuerteAseguradoAdicional = Convert.ToDecimal(row[54].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						SumaAseguradaActualCoberturaITPAseguradoAdicional = Convert.ToDecimal(row[55].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						SobremortalidadCoberturaITPAseguradoAdicional = Convert.ToDecimal(row[56].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						SobreprimaCoberturaITPAseguradoAdicional = Convert.ToDecimal(row[57].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						TasaInteresAcreditadoMes = Convert.ToDecimal(row[58].ToString().Replace(";", "").Replace("%", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						CodigoCorredor = row[59].ToString();
						NombreCorredor = row[60].ToString();


						int _tipoIntermediario = -1;
						switch (row[61].ToString())
						{
							case "Interm. Corredor Persona Juridica":
								_tipoIntermediario = 0;
								break;
							case "Interm. Corredor Persona Natural":
								_tipoIntermediario = 1;
								break;
							case "Intermediario Promotor":
								_tipoIntermediario = 2;
								break;
							case "Personal de la Empresa":
								_tipoIntermediario = 3;
								break;
							case "Trabajador ONSP":
								_tipoIntermediario = 4;
								break;
						}
						TipoIntermediario = _tipoIntermediario;
						IgvComision = row[62].ToString();
						AnioPolizaNumero = Convert.ToInt32(row[63]);
						PorcentajeComisionBrokerPBAnioActual = Convert.ToDecimal(row[64].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						PorcentajeComisionBrokerPEAnioActual = Convert.ToDecimal(row[65].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						ComisionSinIGVPagadoDuranteMes = Convert.ToDecimal(row[66].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						IgvComision2 = (row[67].ToString() == "" ? 0 : Convert.ToDecimal(row[67].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US")));
						ComisionDevengadaDAC = (row[68].ToString() == "" ? 0 : Convert.ToDecimal(row[68].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US")));
						GastoMedicoDevengado = (row[69].ToString() == "" ? 0 : Convert.ToDecimal(row[69].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US")));


						//sw.WriteLine("-------------------------------------------");
						//sw.WriteLine("-------------------------------------------");
						//string valor_prueba = row[70].ToString();
						//sw.WriteLine("ValorPolizaBOP original :: ::  " + valor_prueba.ToString());
						//valor_prueba = valor_prueba.Replace(";", "");
						//sw.WriteLine("ValorPolizaBOP replace ';' :: ::  " + valor_prueba.ToString());
						//valor_prueba = valor_prueba.Replace("%", "");
						//sw.WriteLine("ValorPolizaBOP replace '%' :: ::  " + valor_prueba.ToString());
						//valor_prueba = valor_prueba.Replace(",", "");
						//sw.WriteLine("ValorPolizaBOP replace ',' :: ::  " + valor_prueba.ToString());
						//valor_prueba = valor_prueba.Replace("(", "-");
						//sw.WriteLine("ValorPolizaBOP replace '( por -' :: ::  " + valor_prueba.ToString());
						//valor_prueba = valor_prueba.Replace(")", "-");
						//sw.WriteLine("ValorPolizaBOP replace ')' :: ::  " + valor_prueba.ToString());
						//sw.WriteLine("ValorPolizaBOP final :: ::  " + valor_prueba.ToString());

						//string valor1 = valor_prueba.ToString().Replace(".",",");
						//sw.WriteLine("Valor 1 a Double :: ::  " + Convert.ToDouble(valor1));
						//string valor2 = valor_prueba.ToString().Replace(";",".");
						//sw.WriteLine("Valor 2 a Double :: ::  " + Convert.ToDouble(valor2));

						//string valor3 = valor_prueba.ToString();
						//sw.WriteLine("Valor 3 a Decimal :: ::  " + Convert.ToDecimal(valor3, new CultureInfo("en-US")));

						//sw.WriteLine("ValorPolizaBOP final Convertido a Double :: ::  " + Convert.ToDouble(valor_prueba.ToString()));
						//decimal valor_new = 0;
						//sw.WriteLine("ValorPolizaBOP final Convertido a Decimal :: ::  " + Decimal.TryParse(valor_prueba.ToString(), out valor_new));
						//sw.WriteLine("ValorPolizaBOP final - Decimal :: ::  " + valor_new);
						//sw.WriteLine("ValorPolizaBOP final en otro valor decimal a string :: ::  " + Convert.ToString(valor_new));

						////sw.WriteLine("-------------------------------------------");
						////sw.WriteLine("-------------------------------------------  " + row[70].ToString());
						////sw.WriteLine("Data  :: " + NumeroPoliza + "  ::  " + row[70].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
						//sw.WriteLine("-------------------------------------------");
						//Convert.ToDecimal(valor3, new CultureInfo("en-US")); // 
						ValorPolizaBOP = Convert.ToDecimal(row[70].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						////sw.WriteLine("ValorPolizaBOP :: ::  " + ValorPolizaBOP.ToString());


						//sw.WriteLine("Data  :: " + NumeroPoliza + "  ::  " + row[71].ToString());
						//sw.WriteLine("-------------------------------------------");
						PrimaPagada = Convert.ToDecimal(row[71].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						//sw.WriteLine("PrimaPagada :: ::  " + PrimaPagada.ToString());
						//sw.WriteLine("************************************************");

						//sw.WriteLine("Data  :: " + NumeroPoliza + "  ::  " + row[72].ToString());
						//sw.WriteLine("-------------------------------------------");
						Interes = Convert.ToDecimal(row[72].ToString(), new CultureInfo("en-US"));
						//sw.WriteLine("Interes :: ::  " + Interes.ToString());
						//sw.WriteLine("************************************************");





						RescatesParciales = Convert.ToDecimal(row[73].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						CargoPorRescateParcial = Convert.ToDecimal(row[74].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						GastoAdministrativoPB = Convert.ToDecimal(row[75].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						GastoAdministrativoPexc = Convert.ToDecimal(row[76].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						GastoAdministrativoDisminucionSA = Convert.ToDecimal(row[77].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));

						////sw.WriteLine("-------------------------------------------");
						////sw.WriteLine("-------------------------------------------  " + row[78].ToString());
						////sw.WriteLine("Data  :: " + NumeroPoliza + "  ::  " + row[78].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
						COI_Muerte = Convert.ToDecimal(row[78].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						////sw.WriteLine("COI_Muerte :: ::  " + COI_Muerte.ToString());
						////sw.WriteLine("************************************************");
						////sw.WriteLine("************************************************");

						COI_ITP = Convert.ToDecimal(row[79].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						COI_MuerteAccidental = Convert.ToDecimal(row[80].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						COI_EnfermedadesGraves = Convert.ToDecimal(row[81].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						COI_Exoneracion = Convert.ToDecimal(row[82].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						COI_MuerteAA = Convert.ToDecimal(row[83].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						COI_ITP_AA = Convert.ToDecimal(row[84].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						CargoPorRescateTotalCancelacion = Convert.ToDecimal(row[85].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						COIS_PendientesPago = Convert.ToDecimal(row[86].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						ValorPolizaEOP = Convert.ToDecimal(row[87].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));
						ValorRescate = Convert.ToDecimal(row[88].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""), new CultureInfo("en-US"));



						//72 campos
						DataTable nDT_Parametros;
						nDT_Parametros = nObj.Obtener_Listado("sp_vida_universal_data_reserva_actuarial_operacion", anio_periodo, mes_periodo,
							Moneda, NumeroPoliza, NumeroSolicitud, EstadoPoliza, FechaInicioVigente, FechaCancelacionPoliza, MotivoCancelacion, FechaUltimoEndoso, FechaSuspensionPoliza, TipoSeguro, PeriodicidadPrima, NombreContratante, TipoDocumentoContratante,
		NumeroDocumentoContratante, ContratantePersonaJuridica, Temporalidad, Duracion, AniosPago, MetodoPagoPoliza, PrimaBasica1erAnio,
		PrimaBasicaDemasAnios, PrimaMinima, PrimaMaxima, PrimaPactadaAnual, NombreAseguradoTitular, TipoDocumentoAseguradoTitular,
		NumeroDocumentoAseguradoTitular, FechaNacimientoAseguradoTitular, EdadActuarialAseguradoTitular, SexoAseguradoTitular,
		ClaseRiesgoAseguradoTitular, FumadorAseguradoTitular, SumaAseguradaActualCoberturaMuerte, SobremortalidadCoberturaMuerte,
		SobreprimaCoberturaMuerte, SumaAseguradaActualCoberturaITP, SobremortalidadCoberturaITP, SobreprimaCoberturaITP,
		SumaAseguradaActualCoberturaMuerteAccidental, SobremortalidadCoberturaMuerteAccidental, SobreprimaCoberturaMuerteAccidental,
		SumaAseguradaActualCoberturaEnfermedadesGraves, CubiertoCoberturaExoneracion, NombreAseguradoAdicional,
		TipoDocumentoAseguradoAdicional, NumeroDocumentoAseguradoAdicional, FechaNacimientoAseguradoAdicional, EdadActuarialEmisionAseguradoAdicional,
		SexoAseguradoAdicional, ClaseRiesgoAseguradoAdicional, FumadorAseguradoAdicional, RelacionConTitular,
		SumaAseguradaActualCoberturaMuerteAseguradoAdicional, SobremortalidadCoberturaMuerteAseguradoAdicional, SobreprimaCoberturaMuerteAseguradoAdicional,
		SumaAseguradaActualCoberturaITPAseguradoAdicional, SobremortalidadCoberturaITPAseguradoAdicional, SobreprimaCoberturaITPAseguradoAdicional,
		TasaInteresAcreditadoMes, CodigoCorredor, NombreCorredor, TipoIntermediario, IgvComision, AnioPolizaNumero,
		PorcentajeComisionBrokerPBAnioActual, PorcentajeComisionBrokerPEAnioActual, ComisionSinIGVPagadoDuranteMes, IgvComision2,
		ComisionDevengadaDAC, GastoMedicoDevengado, ValorPolizaBOP,
						PrimaPagada,
						Interes,
						RescatesParciales,
						CargoPorRescateParcial,
						GastoAdministrativoPB,
						GastoAdministrativoPexc,
						GastoAdministrativoDisminucionSA,
						COI_Muerte,
						COI_ITP,
						COI_MuerteAccidental,
						COI_EnfermedadesGraves,
						COI_Exoneracion,
						COI_MuerteAA,
						COI_ITP_AA,
						CargoPorRescateTotalCancelacion,
						COIS_PendientesPago,
						ValorPolizaEOP,
						ValorRescate).Tables[0];

						foreach (DataRow row2 in nDT_Parametros.Rows)
						{
							result = Convert.ToBoolean(row2["result"]);
							mensaje = Convert.ToString(row2["mensaje"]);
						}
						nDT_Parametros = null;

						if (!result)
							break;

					}
					else
					{
						if (posicion >= 6)
							break;
					}

				}

				//Close the file
				////sw.Close();
			}
			catch (Exception ex)
			{
				//ex.Message;
				//Response.Write("Welcome!  ::  " + dt.Rows.Count);
				//Console.WriteLine(ex.Message);                
				//this.Response.Write(" ERROR :: " + ex.Message);
				StreamWriter er = new StreamWriter("C:\\Backup\\Error.txt");
				er.WriteLine(ex.Message);
				er.Close();
			}
			finally
			{
				//oledbConn.Close();
			}

			return result;
		}


		public bool PolizasDatabaseExcel(DataTable dt, string moneda, int anio, int mes, string ruta)
		{
			bool result = true;
			string mensaje = "";

			ModeloNegocio.MyNegocio nObj = new ModeloNegocio.MyNegocio();

			int anio_periodo;
			int mes_periodo;
			string Moneda;
			string NumeroPoliza;
			int NumeroSolicitud;
			int EstadoPoliza;
			DateTime? FechaInicioVigente;
			DateTime? FechaCancelacionPoliza;
			string MotivoCancelacion;
			DateTime? FechaUltimoEndoso;
			DateTime? FechaSuspensionPoliza;
			string TipoSeguro;
			string PeriodicidadPrima;
			string NombreContratante;
			int TipoDocumentoContratante;
			string NumeroDocumentoContratante;
			string ContratantePersonaJuridica;
			int Temporalidad;
			int Duracion;
			int AniosPago;
			string MetodoPagoPoliza;
			decimal PrimaBasica1erAnio;
			decimal PrimaBasicaDemasAnios;
			decimal PrimaMinima;
			decimal PrimaMaxima;
			decimal PrimaPactadaAnual;
			string NombreAseguradoTitular;
			int TipoDocumentoAseguradoTitular;
			string NumeroDocumentoAseguradoTitular;
			DateTime? FechaNacimientoAseguradoTitular;
			int EdadActuarialAseguradoTitular;
			string SexoAseguradoTitular;
			string ClaseRiesgoAseguradoTitular;
			string FumadorAseguradoTitular;
			decimal SumaAseguradaActualCoberturaMuerte;
			decimal SobremortalidadCoberturaMuerte;
			decimal SobreprimaCoberturaMuerte;
			decimal SumaAseguradaActualCoberturaITP;
			decimal SobremortalidadCoberturaITP;
			decimal SobreprimaCoberturaITP;
			decimal SumaAseguradaActualCoberturaMuerteAccidental;
			decimal SobremortalidadCoberturaMuerteAccidental;
			decimal SobreprimaCoberturaMuerteAccidental;
			decimal SumaAseguradaActualCoberturaEnfermedadesGraves;
			string CubiertoCoberturaExoneracion;
			string NombreAseguradoAdicional;
			int TipoDocumentoAseguradoAdicional;
			string NumeroDocumentoAseguradoAdicional;
			DateTime? FechaNacimientoAseguradoAdicional;
			int EdadActuarialEmisionAseguradoAdicional;
			string SexoAseguradoAdicional;
			string ClaseRiesgoAseguradoAdicional;
			string FumadorAseguradoAdicional;
			string RelacionConTitular;
			decimal SumaAseguradaActualCoberturaMuerteAseguradoAdicional;
			decimal SobremortalidadCoberturaMuerteAseguradoAdicional;
			decimal SobreprimaCoberturaMuerteAseguradoAdicional;
			decimal SumaAseguradaActualCoberturaITPAseguradoAdicional;
			decimal SobremortalidadCoberturaITPAseguradoAdicional;
			decimal SobreprimaCoberturaITPAseguradoAdicional;
			decimal TasaInteresAcreditadoMes;
			string CodigoCorredor;
			string NombreCorredor;
			int TipoIntermediario;
			string IgvComision;
			int AnioPolizaNumero;
			decimal PorcentajeComisionBrokerPBAnioActual;
			decimal PorcentajeComisionBrokerPEAnioActual;
			decimal ComisionSinIGVPagadoDuranteMes;
			decimal IgvComision2;
			decimal ComisionDevengadaDAC;
			decimal GastoMedicoDevengado;

			decimal ValorPolizaBOP;
			decimal PrimaPagada;
			decimal Interes;
			decimal RescatesParciales;
			decimal CargoPorRescateParcial;
			decimal GastoAdministrativoPB;
			decimal GastoAdministrativoPexc;
			decimal GastoAdministrativoDisminucionSA;
			decimal COI_Muerte;
			decimal COI_ITP;
			decimal COI_MuerteAccidental;
			decimal COI_EnfermedadesGraves;
			decimal COI_Exoneracion;
			decimal COI_MuerteAA;
			decimal COI_ITP_AA;
			decimal CargoPorRescateTotalCancelacion;
			decimal COIS_PendientesPago;
			decimal ValorPolizaEOP;
			decimal ValorRescate;

			try
			{
				string filePath = ruta;
				int posicion = 1;
				string fecha_null = "01/01/1000";

				using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
				{

					// Auto-detect format, supports:
					//  - Binary Excel files (2.0-2003 format; *.xls)
					//  - OpenXml Excel files (2007 format; *.xlsx)
					using (var reader = ExcelReaderFactory.CreateReader(stream))
					{
						// Choose one of either 1 or 2:
						// 1. Use the reader methods
						do
						{
							while (reader.Read())
							{
								if (reader.RowCount >= posicion)
								{
									if (posicion > 5)
									{

										//reader.GetDouble(0);
										//reader.GetString(1);


										if (reader.GetString(2) == "" || reader.GetString(2) == null)
										{
											break;
										}

										anio_periodo = anio;
										mes_periodo = mes;
										Moneda = moneda;
										NumeroPoliza = (Convert.ToString(reader.GetString(2)) == "" || Convert.ToString(reader.GetString(2)) == null ? "" : Convert.ToString(reader.GetString(2))); //Convert.ToString(row[2]);
										NumeroSolicitud = (Convert.ToString(reader.GetDouble(1)) == "" || Convert.ToString(reader.GetDouble(1)) == null ? 0 : (int)reader.GetDouble(1)); // Convert.ToInt32(row[1]);

										//if (NumeroPoliza == "VIS-UN-000110-0")
										//{
										//	anio_periodo = anio;
										//}

										int _estado = -1;
										if (Convert.ToString(reader.GetString(3)) != "" && Convert.ToString(reader.GetString(3)) != null)
										{
											switch (Convert.ToString(reader.GetString(3)).ToUpper())
											{
												case "ANULADA":
													_estado = 0;
													break;
												case "CANCELACIÓN POR FALTA DE PAGO":
													_estado = 1;
													break;
												case "CANCELACION POR FALTA DE PAGO":
													_estado = 1;
													break;
												case "CANCELACIÓN POR RESCATE":
													_estado = 2;
													break;
												case "CANCELACION POR RESCATE":
													_estado = 2;
													break;
												case "CANCELADO POR RESCATE":
													_estado = 3;
													break;
												case "COBERTURA SUSPENDIDA":
													_estado = 4;
													break;
												case "PERIODO DE GRACIA":
													_estado = 5;
													break;
												case "PÓLIZA VIGENTE":
													_estado = 6;
													break;
												case "POLIZA VIGENTE":
													_estado = 6;
													break;
												case "SOLICITUD APROBADA EN COBRANZA":
													_estado = 7;
													break;
												case "PÓLIZA CANCELADA":
													_estado = 8;
													break;
												case "POLIZA CANCELADA":
													_estado = 8;
													break;
											}
										}
										EstadoPoliza = _estado;

										//FechaInicioVigente = (Convert.ToString(reader.GetString(4)) == "" || Convert.ToString(reader.GetString(4)) == null ? Convert.ToDateTime(fecha_null) : Convert.ToDateTime(reader.GetString(4), new CultureInfo("es-ES")));
										if (!reader.IsDBNull(4))
											FechaInicioVigente = Convert.ToDateTime(reader.GetDateTime(4), new CultureInfo("es-ES")); // (Convert.ToDateTime(reader.GetDateTime(4)) == null ? Convert.ToDateTime(fecha_null) : Convert.ToDateTime(reader.GetDateTime(4), new CultureInfo("es-ES")));
										else
											FechaInicioVigente = Convert.ToDateTime(fecha_null);

										//FechaCancelacionPoliza = (Convert.ToString(reader.GetString(5)) == "" || Convert.ToString(reader.GetString(5)) == null ? Convert.ToDateTime(fecha_null) : Convert.ToDateTime(reader.GetString(5), new CultureInfo("es-ES")));									
										if (!reader.IsDBNull(5))
											FechaCancelacionPoliza = Convert.ToDateTime(reader.GetDateTime(5), new CultureInfo("es-ES")); // (Convert.ToDateTime(reader.GetDateTime(4)) == null ? Convert.ToDateTime(fecha_null) : Convert.ToDateTime(reader.GetDateTime(4), new CultureInfo("es-ES")));
										else
											FechaCancelacionPoliza = Convert.ToDateTime(fecha_null);

										MotivoCancelacion = (Convert.ToString(reader.GetString(6)) == "" || Convert.ToString(reader.GetString(6)) == null ? "" : Convert.ToString(reader.GetString(6)));

										//FechaUltimoEndoso = (Convert.ToString(reader.GetString(7)) == "" || Convert.ToString(reader.GetString(7)) == null ? Convert.ToDateTime(fecha_null) : Convert.ToDateTime(reader.GetString(7), new CultureInfo("es-ES")));									
										if (!reader.IsDBNull(7))
											FechaUltimoEndoso = Convert.ToDateTime(reader.GetDateTime(7), new CultureInfo("es-ES")); // (Convert.ToDateTime(reader.GetDateTime(4)) == null ? Convert.ToDateTime(fecha_null) : Convert.ToDateTime(reader.GetDateTime(4), new CultureInfo("es-ES")));
										else
											FechaUltimoEndoso = Convert.ToDateTime(fecha_null);

										//FechaSuspensionPoliza = (Convert.ToString(reader.GetString(8)) == "" || Convert.ToString(reader.GetString(8)) == null ? Convert.ToDateTime(fecha_null) : Convert.ToDateTime(reader.GetString(8), new CultureInfo("es-ES")));									
										if (!reader.IsDBNull(8))
											FechaSuspensionPoliza = Convert.ToDateTime(reader.GetDateTime(8), new CultureInfo("es-ES")); // (Convert.ToDateTime(reader.GetDateTime(4)) == null ? Convert.ToDateTime(fecha_null) : Convert.ToDateTime(reader.GetDateTime(4), new CultureInfo("es-ES")));
										else
											FechaSuspensionPoliza = Convert.ToDateTime(fecha_null);


										TipoSeguro = (Convert.ToString(reader.GetString(9)) == "" || Convert.ToString(reader.GetString(9)) == null ? "" : Convert.ToString(reader.GetString(9)));

										string _periodicidad = "";
										if (Convert.ToString(reader.GetString(10)) != "" && Convert.ToString(reader.GetString(10)) != null)
										{
											switch (Convert.ToString(reader.GetString(10)))
											{
												case "Anual":
													_periodicidad = "A";
													break;
												case "Mensual":
													_periodicidad = "M";
													break;
												case "Semestral":
													_periodicidad = "S";
													break;
												case "Trimestral":
													_periodicidad = "T";
													break;
											}
										}
										PeriodicidadPrima = _periodicidad;

										NombreContratante = (Convert.ToString(reader.GetString(11)) == "" || Convert.ToString(reader.GetString(11)) == null ? "" : Convert.ToString(reader.GetString(11)));

										int _tipoDocCont = -1;
										if (Convert.ToString(reader.GetString(12)) != "" && Convert.ToString(reader.GetString(12)) != null)
										{
											switch (Convert.ToString(reader.GetString(12)))
											{
												case "DNI":
													_tipoDocCont = 0;
													break;
												case "RUC":
													_tipoDocCont = 1;
													break;
												case "Carnet de extranjeria":
													_tipoDocCont = 2;
													break;
												case "Carnét de extranjería":
													_tipoDocCont = 2;
													break;
											}
										}
										TipoDocumentoContratante = _tipoDocCont;

										NumeroDocumentoContratante = (Convert.ToString(reader.GetString(13)) == "" || Convert.ToString(reader.GetString(13)) == null ? "" : Convert.ToString(reader.GetString(13)));
										ContratantePersonaJuridica = (Convert.ToString(reader.GetString(14)) == "" || Convert.ToString(reader.GetString(14)) == null ? "" : reader.GetString(14).ToString().Substring(0, 1));
										Temporalidad = (Convert.ToString(reader.GetDouble(15)) == "" || Convert.ToString(reader.GetDouble(15)) == null ? 0 : Convert.ToInt32(reader.GetDouble(15)));
										Duracion = (Convert.ToString(reader.GetDouble(16)) == "" || Convert.ToString(reader.GetDouble(16)) == null ? 0 : Convert.ToInt32(reader.GetDouble(16)));
										AniosPago = (Convert.ToString(reader.GetDouble(17)) == "" || Convert.ToString(reader.GetDouble(17)) == null ? 0 : Convert.ToInt32(reader.GetDouble(17)));
										MetodoPagoPoliza = (Convert.ToString(reader.GetString(18)) == "" || Convert.ToString(reader.GetString(18)) == null ? "" : reader.GetString(18).ToString().Substring(0, 1));

										PrimaBasica1erAnio = (Convert.ToString(reader.GetDouble(19)) == "" || Convert.ToString(reader.GetDouble(19)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(19), new CultureInfo("en-US")));
										PrimaBasicaDemasAnios = (Convert.ToString(reader.GetDouble(20)) == "" || Convert.ToString(reader.GetDouble(20)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(20), new CultureInfo("en-US")));
										PrimaMinima = (Convert.ToString(reader.GetDouble(21)) == "" || Convert.ToString(reader.GetDouble(21)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(21), new CultureInfo("en-US")));
										PrimaMaxima = (Convert.ToString(reader.GetDouble(22)) == "" || Convert.ToString(reader.GetDouble(22)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(22), new CultureInfo("en-US")));
										PrimaPactadaAnual = (Convert.ToString(reader.GetDouble(23)) == "" || Convert.ToString(reader.GetDouble(23)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(23), new CultureInfo("en-US")));

										NombreAseguradoTitular = (Convert.ToString(reader.GetString(24)) == "" || Convert.ToString(reader.GetString(24)) == null ? "" : reader.GetString(24).ToString());

										int _tipoDocCont2 = -1;
										if (Convert.ToString(reader.GetString(25)) != "" && Convert.ToString(reader.GetString(25)) != null)
										{
											switch (reader.GetString(25).ToString())
											{
												case "DNI":
													_tipoDocCont2 = 0;
													break;
												case "RUC":
													_tipoDocCont2 = 1;
													break;
												case "Carnet de extranjeria":
													_tipoDocCont2 = 2;
													break;
												case "Carnét de extranjería":
													_tipoDocCont2 = 2;
													break;
											}
										}
										TipoDocumentoAseguradoTitular = _tipoDocCont2;
										NumeroDocumentoAseguradoTitular = (Convert.ToString(reader.GetString(26)) == "" || Convert.ToString(reader.GetString(26)) == null ? "" : reader.GetString(26).ToString());

										//FechaNacimientoAseguradoTitular = (Convert.ToString(reader.GetString(27)) == "" || Convert.ToString(reader.GetString(27)) == null ? Convert.ToDateTime(fecha_null) : Convert.ToDateTime(reader.GetString(27), new CultureInfo("es-ES")));									
										if (!reader.IsDBNull(27))
											FechaNacimientoAseguradoTitular = Convert.ToDateTime(reader.GetDateTime(27), new CultureInfo("es-ES")); // (Convert.ToDateTime(reader.GetDateTime(4)) == null ? Convert.ToDateTime(fecha_null) : Convert.ToDateTime(reader.GetDateTime(4), new CultureInfo("es-ES")));
										else
											FechaNacimientoAseguradoTitular = Convert.ToDateTime(fecha_null);

										EdadActuarialAseguradoTitular = (Convert.ToString(reader.GetDouble(28)) == "" || Convert.ToString(reader.GetDouble(28)) == null ? 0 : Convert.ToInt32(reader.GetDouble(28)));
										SexoAseguradoTitular = (Convert.ToString(reader.GetString(29)) == "" || Convert.ToString(reader.GetString(29)) == null ? "" : reader.GetString(29).ToString().Substring(0, 1));
										ClaseRiesgoAseguradoTitular = (Convert.ToString(reader.GetString(30)) == "" || Convert.ToString(reader.GetString(30)) == null ? "" : reader.GetString(30).ToString().Substring(0, 1));
										FumadorAseguradoTitular = (Convert.ToString(reader.GetString(31)) == "" || Convert.ToString(reader.GetString(31)) == null ? "" : reader.GetString(31).ToString().Substring(0, 1));

										SumaAseguradaActualCoberturaMuerte = (Convert.ToString(reader.GetDouble(32)) == "" || Convert.ToString(reader.GetDouble(32)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(32), new CultureInfo("en-US")));
										SobremortalidadCoberturaMuerte = (Convert.ToString(reader.GetDouble(33)) == "" || Convert.ToString(reader.GetDouble(33)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(33), new CultureInfo("en-US")));
										SobreprimaCoberturaMuerte = (Convert.ToString(reader.GetDouble(34)) == "" || Convert.ToString(reader.GetDouble(34)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(34), new CultureInfo("en-US")));
										SumaAseguradaActualCoberturaITP = (Convert.ToString(reader.GetDouble(35)) == "" || Convert.ToString(reader.GetDouble(35)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(35), new CultureInfo("en-US")));
										SobremortalidadCoberturaITP = (Convert.ToString(reader.GetDouble(36)) == "" || Convert.ToString(reader.GetDouble(36)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(36), new CultureInfo("en-US")));
										SobreprimaCoberturaITP = (Convert.ToString(reader.GetDouble(37)) == "" || Convert.ToString(reader.GetDouble(37)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(37), new CultureInfo("en-US")));
										SumaAseguradaActualCoberturaMuerteAccidental = (Convert.ToString(reader.GetDouble(38)) == "" || Convert.ToString(reader.GetDouble(38)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(38), new CultureInfo("en-US")));
										SobremortalidadCoberturaMuerteAccidental = (Convert.ToString(reader.GetDouble(39)) == "" || Convert.ToString(reader.GetDouble(39)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(39), new CultureInfo("en-US")));
										SobreprimaCoberturaMuerteAccidental = (Convert.ToString(reader.GetDouble(40)) == "" || Convert.ToString(reader.GetDouble(40)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(40), new CultureInfo("en-US")));
										SumaAseguradaActualCoberturaEnfermedadesGraves = (Convert.ToString(reader.GetDouble(41)) == "" || Convert.ToString(reader.GetDouble(41)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(41), new CultureInfo("en-US")));
										CubiertoCoberturaExoneracion = (Convert.ToString(reader.GetString(42)) == "" || Convert.ToString(reader.GetString(42)) == null ? "" : reader.GetString(42).ToString().Substring(0, 1));
										NombreAseguradoAdicional = (Convert.ToString(reader.GetString(43)) == "" || Convert.ToString(reader.GetString(43)) == null ? "" : reader.GetString(43).ToString());

										int _tipoDocCont3 = -1;
										if (Convert.ToString(reader.GetString(44)) != "" && Convert.ToString(reader.GetString(44)) != null)
										{
											switch (reader.GetString(44).ToString())
											{
												case "DNI":
													_tipoDocCont3 = 0;
													break;
												case "RUC":
													_tipoDocCont3 = 1;
													break;
												case "Carnet de extranjeria":
													_tipoDocCont3 = 2;
													break;
												case "Carnét de extranjería":
													_tipoDocCont3 = 2;
													break;
											}
										}
										TipoDocumentoAseguradoAdicional = _tipoDocCont3;
										NumeroDocumentoAseguradoAdicional = (Convert.ToString(reader.GetString(45)) == "" || Convert.ToString(reader.GetString(45)) == null ? "" : reader.GetString(45).ToString());

										//FechaNacimientoAseguradoAdicional = (Convert.ToString(reader.GetString(46)) == "" || Convert.ToString(reader.GetString(46)) == null ? Convert.ToDateTime(fecha_null) : Convert.ToDateTime(reader.GetString(46), new CultureInfo("es-ES")));									
										if (!reader.IsDBNull(46))
											FechaNacimientoAseguradoAdicional = Convert.ToDateTime(reader.GetDateTime(46), new CultureInfo("es-ES")); // (Convert.ToDateTime(reader.GetDateTime(4)) == null ? Convert.ToDateTime(fecha_null) : Convert.ToDateTime(reader.GetDateTime(4), new CultureInfo("es-ES")));
										else
											FechaNacimientoAseguradoAdicional = Convert.ToDateTime(fecha_null);

										EdadActuarialEmisionAseguradoAdicional = (Convert.ToString(reader.GetString(47)) == "" || Convert.ToString(reader.GetString(47)) == null ? 0 : Convert.ToInt32(reader.GetString(47)));  //Convert.ToInt32(row[47]);
										SexoAseguradoAdicional = (Convert.ToString(reader.GetString(48)) == "" || Convert.ToString(reader.GetString(48)) == null ? "" : reader.GetString(48).ToString().Substring(0, 1));  //row[48].ToString().Substring(0, 1);
										ClaseRiesgoAseguradoAdicional = (Convert.ToString(reader.GetString(49)) == "" || Convert.ToString(reader.GetString(49)) == null ? "" : reader.GetString(49).ToString().Substring(0, 1));
										FumadorAseguradoAdicional = (Convert.ToString(reader.GetString(50)) == "" || Convert.ToString(reader.GetString(50)) == null ? "" : reader.GetString(50).ToString().Substring(0, 1));
										RelacionConTitular = (Convert.ToString(reader.GetString(51)) == "" || Convert.ToString(reader.GetString(51)) == null ? "" : reader.GetString(51).ToString());

										SumaAseguradaActualCoberturaMuerteAseguradoAdicional = (Convert.ToString(reader.GetDouble(52)) == "" || Convert.ToString(reader.GetDouble(52)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(52), new CultureInfo("en-US")));
										SobremortalidadCoberturaMuerteAseguradoAdicional = (Convert.ToString(reader.GetDouble(53)) == "" || Convert.ToString(reader.GetDouble(53)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(53), new CultureInfo("en-US")));
										SobreprimaCoberturaMuerteAseguradoAdicional = (Convert.ToString(reader.GetDouble(54)) == "" || Convert.ToString(reader.GetDouble(54)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(54), new CultureInfo("en-US")));
										SumaAseguradaActualCoberturaITPAseguradoAdicional = (Convert.ToString(reader.GetDouble(55)) == "" || Convert.ToString(reader.GetDouble(55)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(55), new CultureInfo("en-US")));
										SobremortalidadCoberturaITPAseguradoAdicional = (Convert.ToString(reader.GetDouble(56)) == "" || Convert.ToString(reader.GetDouble(56)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(56), new CultureInfo("en-US")));
										SobreprimaCoberturaITPAseguradoAdicional = (Convert.ToString(reader.GetDouble(57)) == "" || Convert.ToString(reader.GetDouble(57)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(57), new CultureInfo("en-US")));
										TasaInteresAcreditadoMes = (Convert.ToString(reader.GetDouble(58)) == "" || Convert.ToString(reader.GetDouble(58)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(58), new CultureInfo("en-US")));

										CodigoCorredor = (Convert.ToString(reader.GetString(59)) == "" || Convert.ToString(reader.GetString(59)) == null ? "" : reader.GetString(59).ToString());
										NombreCorredor = (Convert.ToString(reader.GetString(60)) == "" || Convert.ToString(reader.GetString(60)) == null ? "" : reader.GetString(60).ToString());


										int _tipoIntermediario = -1;
										if (Convert.ToString(reader.GetString(61)) != "" && Convert.ToString(reader.GetString(61)) != null)
										{
											switch (reader.GetString(61).ToString())
											{
												case "Interm. Corredor Persona Juridica":
													_tipoIntermediario = 0;
													break;
												case "Interm. Corredor Persona Natural":
													_tipoIntermediario = 1;
													break;
												case "Intermediario Promotor":
													_tipoIntermediario = 2;
													break;
												case "Personal de la Empresa":
													_tipoIntermediario = 3;
													break;
												case "Trabajador ONSP":
													_tipoIntermediario = 4;
													break;
											}
										}
										TipoIntermediario = _tipoIntermediario;
										IgvComision = (Convert.ToString(reader.GetString(62)) == "" || Convert.ToString(reader.GetString(62)) == null ? "" : reader.GetString(62).ToString());
										AnioPolizaNumero = (Convert.ToString(reader.GetDouble(63)) == "" || Convert.ToString(reader.GetDouble(63)) == null ? 0 : Convert.ToInt32(reader.GetDouble(63)));
										PorcentajeComisionBrokerPBAnioActual = (Convert.ToString(reader.GetDouble(64)) == "" || Convert.ToString(reader.GetDouble(64)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(64), new CultureInfo("en-US")));
										PorcentajeComisionBrokerPEAnioActual = (Convert.ToString(reader.GetDouble(65)) == "" || Convert.ToString(reader.GetDouble(65)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(65), new CultureInfo("en-US")));
										ComisionSinIGVPagadoDuranteMes = (Convert.ToString(reader.GetDouble(66)) == "" || Convert.ToString(reader.GetDouble(66)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(66), new CultureInfo("en-US")));

										IgvComision2 = (Convert.ToString(reader.GetString(67)) == "" || Convert.ToString(reader.GetString(67)) == null ? 0 : Convert.ToDecimal(reader.GetString(67), new CultureInfo("en-US")));
										ComisionDevengadaDAC = (Convert.ToString(reader.GetString(68)) == "" || Convert.ToString(reader.GetString(68)) == null ? 0 : Convert.ToDecimal(reader.GetString(68), new CultureInfo("en-US")));
										GastoMedicoDevengado = (Convert.ToString(reader.GetDouble(69)) == "" || Convert.ToString(reader.GetDouble(69)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(69), new CultureInfo("en-US")));

										//Convert.ToDecimal(valor3, new CultureInfo("en-US")); // 
										ValorPolizaBOP = (Convert.ToString(reader.GetDouble(70)) == "" || Convert.ToString(reader.GetDouble(70)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(70), new CultureInfo("en-US")));
										////sw.WriteLine("ValorPolizaBOP :: ::  " + ValorPolizaBOP.ToString());


										//sw.WriteLine("Data  :: " + NumeroPoliza + "  ::  " + row[71].ToString());
										//sw.WriteLine("-------------------------------------------");
										PrimaPagada = (Convert.ToString(reader.GetDouble(71)) == "" || Convert.ToString(reader.GetDouble(71)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(71), new CultureInfo("en-US")));
										//sw.WriteLine("PrimaPagada :: ::  " + PrimaPagada.ToString());
										//sw.WriteLine("************************************************");

										//sw.WriteLine("Data  :: " + NumeroPoliza + "  ::  " + row[72].ToString());
										//sw.WriteLine("-------------------------------------------");
										Interes = (Convert.ToString(reader.GetDouble(72)) == "" || Convert.ToString(reader.GetDouble(72)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(72), new CultureInfo("en-US")));
										//sw.WriteLine("Interes :: ::  " + Interes.ToString());
										//sw.WriteLine("************************************************");

										RescatesParciales = (Convert.ToString(reader.GetDouble(73)) == "" || Convert.ToString(reader.GetDouble(73)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(73), new CultureInfo("en-US")));
										CargoPorRescateParcial = (Convert.ToString(reader.GetDouble(74)) == "" || Convert.ToString(reader.GetDouble(74)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(74), new CultureInfo("en-US")));
										GastoAdministrativoPB = (Convert.ToString(reader.GetDouble(75)) == "" || Convert.ToString(reader.GetDouble(75)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(75), new CultureInfo("en-US")));
										GastoAdministrativoPexc = (Convert.ToString(reader.GetDouble(76)) == "" || Convert.ToString(reader.GetDouble(76)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(76), new CultureInfo("en-US")));
										GastoAdministrativoDisminucionSA = (Convert.ToString(reader.GetDouble(77)) == "" || Convert.ToString(reader.GetDouble(77)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(77), new CultureInfo("en-US")));

										////sw.WriteLine("-------------------------------------------");
										////sw.WriteLine("-------------------------------------------  " + row[78].ToString());
										////sw.WriteLine("Data  :: " + NumeroPoliza + "  ::  " + row[78].ToString().Replace(";", "").Replace("%", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
										COI_Muerte = (Convert.ToString(reader.GetDouble(78)) == "" || Convert.ToString(reader.GetDouble(78)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(78), new CultureInfo("en-US")));
										////sw.WriteLine("COI_Muerte :: ::  " + COI_Muerte.ToString());
										////sw.WriteLine("************************************************");
										////sw.WriteLine("************************************************");

										COI_ITP = (Convert.ToString(reader.GetDouble(79)) == "" || Convert.ToString(reader.GetDouble(79)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(79), new CultureInfo("en-US")));
										COI_MuerteAccidental = (Convert.ToString(reader.GetDouble(80)) == "" || Convert.ToString(reader.GetDouble(80)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(80), new CultureInfo("en-US")));
										COI_EnfermedadesGraves = (Convert.ToString(reader.GetDouble(81)) == "" || Convert.ToString(reader.GetDouble(81)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(81), new CultureInfo("en-US")));
										COI_Exoneracion = (Convert.ToString(reader.GetDouble(82)) == "" || Convert.ToString(reader.GetDouble(82)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(82), new CultureInfo("en-US")));
										COI_MuerteAA = (Convert.ToString(reader.GetDouble(83)) == "" || Convert.ToString(reader.GetDouble(83)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(83), new CultureInfo("en-US")));
										COI_ITP_AA = (Convert.ToString(reader.GetDouble(84)) == "" || Convert.ToString(reader.GetDouble(84)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(84), new CultureInfo("en-US")));
										CargoPorRescateTotalCancelacion = (Convert.ToString(reader.GetDouble(85)) == "" || Convert.ToString(reader.GetDouble(85)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(85), new CultureInfo("en-US")));
										COIS_PendientesPago = (Convert.ToString(reader.GetDouble(86)) == "" || Convert.ToString(reader.GetDouble(86)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(86), new CultureInfo("en-US")));
										ValorPolizaEOP = (Convert.ToString(reader.GetDouble(87)) == "" || Convert.ToString(reader.GetDouble(87)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(87), new CultureInfo("en-US")));
										ValorRescate = (Convert.ToString(reader.GetDouble(88)) == "" || Convert.ToString(reader.GetDouble(88)) == null ? 0 : Convert.ToDecimal(reader.GetDouble(88), new CultureInfo("en-US")));



										//72 campos
										DataTable nDT_Parametros;
										nDT_Parametros = nObj.Obtener_Listado("SP_VIDA_TEMPORAL_DATA_RESERVA_ACTUARIAL_OPERACION", anio_periodo, mes_periodo,
											Moneda, NumeroPoliza, NumeroSolicitud, EstadoPoliza, FechaInicioVigente, FechaCancelacionPoliza, MotivoCancelacion, FechaUltimoEndoso, FechaSuspensionPoliza, TipoSeguro, PeriodicidadPrima, NombreContratante, TipoDocumentoContratante,
						NumeroDocumentoContratante, ContratantePersonaJuridica, Temporalidad, Duracion, AniosPago, MetodoPagoPoliza, PrimaBasica1erAnio,
						PrimaBasicaDemasAnios, PrimaMinima, PrimaMaxima, PrimaPactadaAnual, NombreAseguradoTitular, TipoDocumentoAseguradoTitular,
						NumeroDocumentoAseguradoTitular, FechaNacimientoAseguradoTitular, EdadActuarialAseguradoTitular, SexoAseguradoTitular,
						ClaseRiesgoAseguradoTitular, FumadorAseguradoTitular, SumaAseguradaActualCoberturaMuerte, SobremortalidadCoberturaMuerte,
						SobreprimaCoberturaMuerte, SumaAseguradaActualCoberturaITP, SobremortalidadCoberturaITP, SobreprimaCoberturaITP,
						SumaAseguradaActualCoberturaMuerteAccidental, SobremortalidadCoberturaMuerteAccidental, SobreprimaCoberturaMuerteAccidental,
						SumaAseguradaActualCoberturaEnfermedadesGraves, CubiertoCoberturaExoneracion, NombreAseguradoAdicional,
						TipoDocumentoAseguradoAdicional, NumeroDocumentoAseguradoAdicional, FechaNacimientoAseguradoAdicional, EdadActuarialEmisionAseguradoAdicional,
						SexoAseguradoAdicional, ClaseRiesgoAseguradoAdicional, FumadorAseguradoAdicional, RelacionConTitular,
						SumaAseguradaActualCoberturaMuerteAseguradoAdicional, SobremortalidadCoberturaMuerteAseguradoAdicional, SobreprimaCoberturaMuerteAseguradoAdicional,
						SumaAseguradaActualCoberturaITPAseguradoAdicional, SobremortalidadCoberturaITPAseguradoAdicional, SobreprimaCoberturaITPAseguradoAdicional,
						TasaInteresAcreditadoMes, CodigoCorredor, NombreCorredor, TipoIntermediario, IgvComision, AnioPolizaNumero,
						PorcentajeComisionBrokerPBAnioActual, PorcentajeComisionBrokerPEAnioActual, ComisionSinIGVPagadoDuranteMes, IgvComision2,
						ComisionDevengadaDAC, GastoMedicoDevengado, ValorPolizaBOP,
										PrimaPagada,
										Interes,
										RescatesParciales,
										CargoPorRescateParcial,
										GastoAdministrativoPB,
										GastoAdministrativoPexc,
										GastoAdministrativoDisminucionSA,
										COI_Muerte,
										COI_ITP,
										COI_MuerteAccidental,
										COI_EnfermedadesGraves,
										COI_Exoneracion,
										COI_MuerteAA,
										COI_ITP_AA,
										CargoPorRescateTotalCancelacion,
										COIS_PendientesPago,
										ValorPolizaEOP,
										ValorRescate).Tables[0];

										foreach (DataRow row2 in nDT_Parametros.Rows)
										{
											result = Convert.ToBoolean(row2["result"]);
											mensaje = Convert.ToString(row2["mensaje"]);
										}
										nDT_Parametros = null;

										if (!result)
											break;

									}
									posicion += 1;
								}
							}
						} while (reader.NextResult());

						// 2. Use the AsDataSet extension method
						//var result = reader.AsDataSet();

						// The result of each spreadsheet is in result.Tables
					}
				}



			}
			catch (Exception ex)
			{
				//ex.Message;
				result = false;
				//Response.Write("Welcome!  ::  " + dt.Rows.Count);
				//Console.WriteLine(ex.Message);                
				//this.Response.Write(" ERROR :: " + ex.Message);

				//StreamWriter er = new StreamWriter("C:\\Backup\\Error.txt");
				//er.WriteLine(ex.Message);
				//er.Close();
			}
			finally
			{
				//oledbConn.Close();
			}

			return result;
		}
	}
}