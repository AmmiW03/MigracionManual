using MongoDB.Bson;
using MongoExcelMigration.Modelos;
using NPOI.POIFS.Crypt.Dsig;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using SixLabors.ImageSharp.Processing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MongoExcelMigration.Modelos
{
    //Versión: 1.0
    //Fecha: 04 de Abril de 2023
    //Autor: ASLOGIC S.A. DE C.V.
    //Desarrollador: Ammi Jatziry Wang Almazán
    //Módulo: Modelo de Métodos.
    //Descripción: Clase con los métodos y procesos utilizados apra su migración a MongoDB.
    //Historial de cambios:
    //04 de Abril de 2023: Se soluciona el error de la clave única en los encabezados.
    public static class mdlMetodos
    {
        public static void ReadExcel(String filePath, int rowHeaderC, String dbName)
        {
            try
            {
                string fileName = @"D:\Certificacion07-1\Excel\" + filePath;

                FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);

                IWorkbook workbook = new XSSFWorkbook(fs);

                ISheet sheet = workbook.GetSheetAt(0);

                DataFormatter dataf = new();

                List<String> lstHeaders = new List<string>();


                if (sheet == null) return;
                
                IRow headersRow = sheet.GetRow(rowHeaderC);
                int auxInt = 0;
                while (dataf.FormatCellValue(headersRow.GetCell(auxInt)) != "")
                {

                    String compar = dataf.FormatCellValue(headersRow.GetCell(auxInt));
                    if (lstHeaders.Contains(compar))
                    {
                        lstHeaders.Add(compar + auxInt);
                    }
                    else
                    {
                        lstHeaders.Add(compar);
                    }
                    auxInt++;
                }

                List<CellType> cellTypes = new List<CellType>();
                bool flag = true;

                #region Creacion de nombres de campos

                for (int i = 0; i < lstHeaders.Count; i++)
                {
                    switch (lstHeaders[i])
                    {
                        #region Modulo 1

                        #region Columna 1
                        case "No."://Interpretación
                            lstHeaders.Add("em_numero");//bd
                            break;

                        case "Nombre":
                            lstHeaders.Add("em_nombre");
                            break;

                        case "Compañia":
                            lstHeaders.Add("em_cia");
                            break;

                        case "Fecha Ingreso":
                            lstHeaders.Add("em_fechai");
                            break;

                        case "Fecha Antiguedad":
                            lstHeaders.Add("em_fechant");
                            break;

                        case "Fecha Baja":
                            lstHeaders.Add("em_fechab");
                            break;

                        case "U.Camb Sal":
                            lstHeaders.Add("em_fechcam");
                            break;

                        case "Fecha Planta":
                            lstHeaders.Add("em_fecplan");
                            break;

                        case "Fecha U.Cont.":
                            lstHeaders.Add("em_feculco");
                            break;

                        case "E. Civ":
                            lstHeaders.Add("em_estciv");
                            break;

                        case "RFC":
                            lstHeaders.Add("em_rfc");
                            break;

                        case "No.Alfil IMSS":
                            lstHeaders.Add("em_imss");
                            break;

                        case "Gpo. IMSS":
                            lstHeaders.Add("em_gruimss");
                            break;

                        case "Tip Col":
                            lstHeaders.Add("em_tipoemp");
                            break;

                        case "Tip Nom":
                            lstHeaders.Add("em_tiponom");
                            break;

                        #endregion

                        #region Columna 2

                        case "Puesto":
                            lstHeaders.Add("em_puesto");
                            break;

                        case "No, Div":
                            lstHeaders.Add("em_divisio");
                            break;

                        case "Centro de Costo":
                            lstHeaders.Add("em_depto");
                            break;

                        case "C. Pago":
                            lstHeaders.Add("em_cpago");
                            break;

                        case "Turno":
                            lstHeaders.Add("em_turno");
                            break;

                        case "S.Diario":
                            lstHeaders.Add("em_saldia");
                            break;

                        case "S.Propor":
                            lstHeaders.Add("em_salprop");
                            break;

                        case "S.Prom.Prop":
                            lstHeaders.Add("em_salppro");
                            break;

                        case "S.Prom":
                            lstHeaders.Add("em_salprom");
                            break;

                        case "Salario":
                            lstHeaders.Add("em_salario");
                            break;

                        case "T. Sal":
                            lstHeaders.Add("em_tiposal");
                            break;

                        case "S.D.I":
                            lstHeaders.Add("em_salinte");
                            break;

                        case "S.D.I. Var":
                            lstHeaders.Add("em_sdivar");
                            break;

                        case "S.D.I. Ant":
                            lstHeaders.Add("em_asalint");
                            break;

                        case "S.D.I. Var Ant":
                            lstHeaders.Add("em_avarant");
                            break;
                        #endregion

                        #region Columna 3

                        case "Sal. Ant":
                            lstHeaders.Add("em_cambios");
                            break;

                        case "Sexo":
                            lstHeaders.Add("em_sexo");
                            break;

                        case "F.Nac":
                            lstHeaders.Add("em_fechnac");
                            break;

                        case "Z Ec.":
                            lstHeaders.Add("em_ubzona");
                            break;

                        case "Suc":
                            lstHeaders.Add("em_sucursa");
                            break;

                        case "M.O":
                            lstHeaders.Add("em_manobra");
                            break;

                        case "T. San":
                            lstHeaders.Add("em_tiposan");
                            break;

                        case "Tipo Cont":
                            lstHeaders.Add("em_contra");
                            break;

                        case "Nivel":
                            lstHeaders.Add("em_nivel");
                            break;

                        case "Reing":
                            lstHeaders.Add("em_reingre");
                            break;

                        case "Tab":
                            lstHeaders.Add("em_tabula");
                            break;

                        case "Sup":
                            lstHeaders.Add("em_super");
                            break;

                        case "S. Garantia":
                            lstHeaders.Add("em_salgara");
                            break;

                        case "CURP":
                            lstHeaders.Add("em_curp");
                            break;

                        case "Cel":
                            lstHeaders.Add("em_celula");
                            break;
                        #endregion

                        #region Columna 4

                        case "Gpo":
                            lstHeaders.Add("em_grupo");
                            break;

                        case "Sub":
                            lstHeaders.Add("em_subgrp");
                            break;

                        case "S.Tabulado":
                            lstHeaders.Add("m_saltab");
                            break;

                        case "Incentivo":
                            lstHeaders.Add("em_incenti");
                            break;

                        case "Dia Eco":
                            lstHeaders.Add("em_diaeco");
                            break;

                        case "T. Col Ant":
                            lstHeaders.Add("em_tempant");
                            break;

                        case "T. Nom Ant":
                            lstHeaders.Add("em_tnomant");
                            break;

                        case "F. Camb. T.Col":
                            lstHeaders.Add("em_fnewnom");
                            break;

                        case "F.Matrimonio":
                            lstHeaders.Add("em_fecmatr");
                            break;
                        #endregion

                        #endregion

                        #region Modulo 2

                        #region Columna 1

                        case "C. ISPT":
                            lstHeaders.Add("em_cispt");
                            break;

                        case "Ajus Anu.":
                            lstHeaders.Add("em_cajuste");
                            break;

                        case "C. IMSS":
                            lstHeaders.Add("em_cimss");
                            break;

                        case "Pag C.S.":
                            lstHeaders.Add("em_cfisica");
                            break;

                        case "C.Rep PTU":
                            //Pendiente
                            lstHeaders.Add("");
                            break;

                        case "C. Aguin":
                            lstHeaders.Add("em_caguina");
                            break;

                        case "D. Aguin":
                            lstHeaders.Add("em_cdias");
                            break;

                        case "C.V. Desp":
                            lstHeaders.Add("em_valesde");
                            break;

                        case "C.V. Com":
                            lstHeaders.Add("em_valesco");
                            break;

                        case "C.F. Aho.":
                            lstHeaders.Add("em_cfahorr");
                            break;

                        case "C.P. Asis.":
                            lstHeaders.Add("em_asisten");
                            break;

                        case "Imp Rec":
                            lstHeaders.Add("em_irecibo");
                            break;

                        case "Imp Cheq.":
                            lstHeaders.Add("em_iforma");
                            break;

                        case "Abon Ban":
                            lstHeaders.Add("em_abonar");
                            break;

                        case "T. Ban":
                            lstHeaders.Add("em_banco");
                            break;
                        #endregion

                        #region Columna 2

                        case "Suc Ban":
                            lstHeaders.Add("em_sucurba");
                            break;

                        case "Plaza Ban.E":
                            lstHeaders.Add("em_plaza");
                            break;

                        case "Tipo Cta":
                            lstHeaders.Add("em_tipocta");
                            break;

                        case "Cuenta Banco":
                            lstHeaders.Add("em_cuenta");
                            break;

                        case "C. INFO":
                            lstHeaders.Add("em_cinfona");
                            break;

                        case "Cuenta. INFONAVIT":
                            lstHeaders.Add("em_infocre");
                            break;

                        case "% Dcto. Cred.IFONAVIT":
                            lstHeaders.Add("em_infopor");
                            break;

                        case "F.I.C. INFONAVIT":
                            lstHeaders.Add("em_fcreinf");
                            break;

                        case "T.Des. INFONAVIT":
                            lstHeaders.Add("em_tipoinf");
                            break;

                        case "% Pen Alim.":
                            lstHeaders.Add("em_penspor");
                            break;

                        case "Importe Pen Alim":
                            lstHeaders.Add("em_pensimp");
                            break;

                        case "P.V. Aut":
                            lstHeaders.Add("em_porprim");
                            break;

                        //Pendiente
                        case "Ind. Vac":
                            lstHeaders.Add("");
                            break;

                        case "P.Re Vac":
                            lstHeaders.Add("em_pereva");
                            break;

                        case "Ini P.Vac":
                            lstHeaders.Add("em_inpeva");
                            break;

                        case "F.Ini.Vac.":
                            lstHeaders.Add("em_fechaiv");
                            break;

                        #endregion

                        #region Columna 3

                        case "F.Reg.Vac.":
                            lstHeaders.Add("em_retvac");
                            break;

                        case "SDI para 25 SMDF art 33 del SS":
                            lstHeaders.Add("em_sdia29");
                            break;

                        case "SDI para 15 SDMF art 33 del SS":
                            lstHeaders.Add("em_sdib29");
                            break;

                        case "% o cant. a pagar anticipo":
                            lstHeaders.Add("em_anticip");
                            break;

                        case "Cant. de Anti.Sem":
                            lstHeaders.Add("em_antisem");
                            break;

                        case "% Bono":
                            lstHeaders.Add("em_porbono");
                            break;

                        case "Factor Propor":
                            lstHeaders.Add("em_propor");
                            break;

                        case "Fac.C. Sal.Diario":
                            lstHeaders.Add("em_minimom");
                            break;

                        case "Reloj Chec.":
                            lstHeaders.Add("em_reloj");
                            break;

                        case "Act":
                            lstHeaders.Add("em_activi");
                            break;

                        case "Cuenta Contable":
                            lstHeaders.Add("em_ctacont");
                            break;

                        case "% Dcto. por Mant.":
                            lstHeaders.Add("em_infoman");
                            break;

                        case "Asimil. Salario":
                            lstHeaders.Add("em_asimila");
                            break;

                        case "UMF":
                            lstHeaders.Add("em_imssumf");
                            break;

                        case "e-Mail e-Mail":
                            lstHeaders.Add("em_email");
                            break;
                        #endregion

                        #endregion

                        #region Modulo 3

                        #region Columna 1

                        case "Tel.  ":
                            lstHeaders.Add("rh_telefo");
                            break;

                        case "Escolaridad":
                            lstHeaders.Add("rh_escolar");
                            break;

                        case "Cd. Nac.":
                            lstHeaders.Add("rh_nciudad");
                            break;

                        case "Edo. Nac.":
                            lstHeaders.Add("rh_nestado");
                            break;

                        case "Calle":
                            lstHeaders.Add("rh_dcalle");
                            break;

                        case "Colonia":
                            lstHeaders.Add("rh_dcolon");
                            break;

                        case "Ciudad":
                            lstHeaders.Add("rh_dciudad");
                            break;

                        case "Estado":
                            lstHeaders.Add("rh_destado");
                            break;

                        case "Municipio":
                            lstHeaders.Add("rh_dmunici");
                            break;

                        case "C.P.":
                            lstHeaders.Add("rh_dcp");
                            break;

                        case "Nombre del padre":
                            lstHeaders.Add("rh_npadre");
                            break;

                        case "Nombre del madre":
                            lstHeaders.Add("rh_nmadre");
                            break;

                        case "P Fin":
                            lstHeaders.Add("rh_fpadre");
                            break;

                        case "M Fin":
                            lstHeaders.Add("rh_fmadre");
                            break;

                        case "Nacionalidad":
                            lstHeaders.Add("rh_nacion");
                            break;
                        #endregion

                        #region Columna 2

                        case "1er. Asegurado GMM":
                            lstHeaders.Add("rh_gmmaseg");
                            break;

                        case "F.Nac. Asegurado":
                            lstHeaders.Add("rh_gmmfnac");
                            break;

                        case "Sexo":
                            lstHeaders.Add("rh_gmmsexo");
                            break;

                        case "Paren":
                            lstHeaders.Add("rh_gmmpare");
                            break;

                        case "Nom. p/repor emergencia":
                            lstHeaders.Add("rh_gmmpor");
                            break;

                        case "Primer Nombre p/emergencia":
                            lstHeaders.Add("rh_noavis1");
                            break;

                        case "Tel.del emergente":
                            lstHeaders.Add("rh_teavis1");
                            break;

                        case "Parentesco":
                            lstHeaders.Add("rh_paavis1");
                            break;

                        case "Segundo Nombre de emergencia":
                            lstHeaders.Add("rh_noavis2");
                            break;

                        case "Tel del emergente":
                            lstHeaders.Add("rh_teavis2");
                            break;

                        case "Parentesco":
                            lstHeaders.Add("rh_paavis2");
                            break;

                        case "Fotografia Asoc.":
                            lstHeaders.Add("rh_picture");
                            break;

                        case "Cve. P.GMM":
                            lstHeaders.Add("rh_gmmpcve");
                            break;

                        case "Area Col":
                            lstHeaders.Add("rh_area");
                            break;

                        case "Oficio":
                            lstHeaders.Add("rh_oficio");
                            break;
                        #endregion

                        #region Columna 3

                        case "Estat":
                            lstHeaders.Add("be_estatur");
                            break;

                        case "Peso":
                            lstHeaders.Add("be_peso");
                            break;

                        case "GMM":
                            lstHeaders.Add("");
                            break;

                        case "Seg Vida":
                            lstHeaders.Add("");
                            break;

                        case "Suma  Aseg. GMM":
                            lstHeaders.Add("rh_gmmsuma");
                            break;

                        case "Plan Seg.  Vida":
                            lstHeaders.Add("rh_plansv");
                            break;

                        case "P.A. S.V.":
                            lstHeaders.Add("rh_aplansv");
                            break;

                        case "Prima Aseg.S.V.":
                            lstHeaders.Add("rh_psvsuma");
                            break;

                        case "Ubicación del Colaborador":
                            lstHeaders.Add("rh_ubicado");
                            break;

                        case "Estat.":
                            lstHeaders.Add("rh_estatu");
                            break;

                        case "Peso":
                            lstHeaders.Add("rh_peso");

                            break;

                        case "Talla Camisa":
                            lstHeaders.Add("rh_tallac");
                            break;

                        case "Talla Pantalon":
                            lstHeaders.Add("rh_tallap");
                            break;

                        case "Calzado":
                            lstHeaders.Add("rh_calzado");
                            break;

                        case "Color Ojos":
                            lstHeaders.Add("rh_coloroj");
                            break;
                        #endregion

                        #region Columna 4

                        case "Color Cabello":
                            lstHeaders.Add("rh_colorca");
                            break;

                        case "Color Piel":
                            lstHeaders.Add("rh_piel");
                            break;

                        case "Señas Particulares":
                            lstHeaders.Add("rh_separt");
                            break;

                        case "Estudio Soc-Eco":
                            lstHeaders.Add("rh_soceco");
                            break;

                        case "Alta Seg. Pub.":
                            lstHeaders.Add("rh_segpub");
                            break;

                        case "Examen Antidoping":
                            lstHeaders.Add("rh_antidop");
                            break;
                            #endregion

                            #endregion
                    }
                }

                #endregion

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}


//BsonDocument document = new BsonDocument();
//mdlMongoDB.SubirDatos(dbName, data.ToBsonDocument());
