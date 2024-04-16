import pandas as pd
import snowflake.connector
import openpyxl
import datetime
import teradata
import re
import timeit
import getpass
import warnings
warnings.filterwarnings("ignore")

if __name__ == '__main__':
    try:
        start = timeit.default_timer()

        username = input('Nike Username: ')
        password = getpass.getpass('Password: ')

        credentials = {
            "account" :"nike",
            "user" : username,
            "password" :password,
            # "database" :"DA_DSM_SCANALYTICS_PROD",
            "database" :"DA_DSM_SCANALYTICS_PROD",
            "warehouse" :"DA_DSM_SCANALYTICS_REPORTING_PROD",
            "schema" : "PROCESSED",
            # "role":"MPAA_ROLE_PROD_ANALYTICS",
            "role":"Application_SnowFlake_Prod_DSM_SCANALYTICS_Read",
            "auth": "externalbrowser",
            "authenticator" : "https://nike.okta.com"
        }

        #JSON Linked file to setup username, password, and other database details
        def create_sql_engine(credentials):

            # engine = create_engine(URL(
            engine = snowflake.connector.connect(

            account=credentials['account'],
            user=credentials['user'],
            password=credentials['password'],
            database=credentials['database'],
            warehouse=credentials['warehouse'],
            schema=credentials['schema'],
            role=credentials['role'],
            authenticator=credentials['auth'],

            )
            return engine
    
        # engine = create_sql_engine("credentials.json")
        engine = create_sql_engine(credentials)

        # SQL queries

# https://confluence.nike.com/display/EDAEDP/Global+Product+Line

        gel_query = """

        SELECT
            SEASON_YEAR_DESCRIPTION AS seasn_yr_cd,
            DIVISION_CODE AS DIV_CD,
            STYLE_NUMBER AS styl_dsply_cd,
            DEVELOPMENT_TEAM_NAME

        FROM EDA_PRODUCT_PROD.BCL_PRODUCT_LINE.GLOBAL_PRODUCT_LINE_V1

        WHERE 
        DIVISION_CODE IN (10, 20)
        and SEASON_YEAR_DESCRIPTION IN {}
        and (DEVELOPMENT_TEAM_NAME like '%GEL%' or DEVELOPMENT_TEAM_NAME like '%GLOBAL EXPRESS%')
        and STYLE_NUMBER is not null

        GROUP BY 
        SEASON_YEAR_DESCRIPTION, DIVISION_CODE, STYLE_NUMBER, DEVELOPMENT_TEAM_NAME

                """

        biz_season_query = """

        select 
            BUSALTSEASNCD
            ,BUSSEASNSTRTDT
            ,BUSSEASNENDDT
            ,BUSSEASNSORTSEQNBR
        from NGP_DA_PROD.PLN.CAL_BUSFSCLN1GREGORIANCALDT_V
        where BUSCALCD = 'NF'
        group by 1,2,3,4
        order by BUSSEASNSORTSEQNBR
                
                """

        max_engine_run_dt_query = """
        SELECT 
            max(REF_SNP_ENGN_RUN_DT)
        FROM DA_DSM_SCANALYTICS_PROD.INTEGRATED.UNPLANNED_T

                """
        
        #SQL Engine to read the four queries, run one at a time if on local machine
        sf = engine.cursor()
        seasn_query = sf.execute(biz_season_query)
        df_season = pd.DataFrame(seasn_query.fetchall())

        today = datetime.date.today()

        def cur_date(date_today):
            return df_season[(date_today >= df_season[1]) & (date_today <= df_season[2])]

        def seasons(date_today):
            cur_season = cur_date(date_today)
            cur_season_val = cur_season[3].values[0]
            result_df = df_season[(df_season[3] > cur_season_val) & (df_season[3] <= cur_season_val + 4)]
            result_list = result_df[0].values.tolist()
            t = tuple(result_list)
            return gel_query.format(t)

        print('Querying Snowflake for GEL styles...')

        sf_gel_query = sf.execute(seasons(today))
        df_gel = pd.DataFrame(sf_gel_query.fetchall())

        ap_headers = ['PPPriorityID','Plant','DemandSeason','CategoryDesc','SubCategoryDesc','LeagueID','LeagueDesc','StyleCode','ColorCode','PriorityDesc','Reason','RequestedBy','Priority','DefaultPriority','updFlag','Error']

        fw_headers = ['PPPriorityID','Plant','DemandSeason','CategoryDesc','SubCategoryDesc','StyleCode','ColorCode','Reason','RequestedBy','PriorityDesc','Priority','DefaultPriority','updFlag','Error']

        gel_ap_template = pd.DataFrame(columns=ap_headers)
        gel_fw_template = pd.DataFrame(columns=fw_headers)

        df_gel_ap = df_gel[df_gel[1] == '10']
        df_gel_fw = df_gel[df_gel[1] == '20']

        gel_ap_template['DemandSeason'] = df_gel_ap[0]
        gel_ap_template['StyleCode'] = df_gel_ap[2]
        gel_ap_template['PriorityDesc'] = 'P'
        gel_ap_template['Reason'] = 'GEL'
        gel_ap_template['RequestedBy'] = 'GOVERNANCE STANDARD'
        gel_ap_template['Priority'] = '50'
        gel_ap_template['DefaultPriority'] = '100'
        gel_ap_template['updFlag'] = "I"

        gel_fw_template['DemandSeason'] = df_gel_fw[0]
        gel_fw_template['StyleCode'] = df_gel_fw[2]
        gel_fw_template['PriorityDesc'] = 'P'
        gel_fw_template['Reason'] = 'GEL'
        gel_fw_template['RequestedBy'] = 'GOVERNANCE STANDARD'
        gel_fw_template['Priority'] = '50'
        gel_fw_template['DefaultPriority'] = '100'
        gel_fw_template['updFlag'] = "I"

        mer_query = sf.execute(max_engine_run_dt_query)
        max_engine_run_dt = pd.DataFrame(mer_query.fetchall())
        max_engine_run_dt[0] = max_engine_run_dt[0].astype(str)
        max_engine_run_dt = max_engine_run_dt[0].values[0]


        udaExec = teradata.UdaExec(appName="testconnec", version="1.0", logConsole=False)
        session = udaExec.connect(method="odbc", dsn="Teradata_Production", username=username, password=password, authentication="LDAP", driver="TeraProd")

        flyknit_query = """

        select 
        "Cal_SNP_EngnRunDt_v__Version_"."RefSNP_EngnRunDt" "Ver_Wk_Dt" , 
        "Div_v"."DivCd" "Prod_Engn_Cd" , 
        right("T3"."BusSeasnCd", 2) || left("T3"."BusSeasnCd", 4) "SesnYr" , 
        "LnPln_Rptg_v"."StylDsplyCd" "Styl_Dsply_Cd" , 
        "LnPln_Rptg_v"."GlblMatlIntntDesc" "Matl_Intnt_Desc" , 
        "SNP_EnhcdDmndPrty_v"."EnhcdDmndPrtyDesc" "EDP" , 
        "LnPln_Rptg_v"."GlblProdtLglLongNm" "APO_Prod_Desc" , 
        "LnPln_Rptg_v"."MatlTypeCd" "Matl_Type_Cd" , 
        "LnPln_Rptg_v"."PO_DevTeamNm" "PO_Dev_Team_Nm" , 
        "LnPln_Rptg_v"."RgnSeasnSpclOfrgTypeDesc" "c10" , 
        sum("SNP_FinGdCmphsvUnPlndDtl_v"."TotDmndQty") "Tot_Rlsd_DP_Qty" , 
        sum("SNP_FinGdCmphsvUnPlndDtl_v"."NetRqmntsQty") "Net_Rqrmnts_Qty" , 
        sum("SNP_FinGdCmphsvUnPlndDtl_v"."TotUnPlndDmndQty") "Tot_Unplanned_Dmnd_Qty" 
        
        from ((((
            "PLN"."Cal_SNP_EngnRunDt_v" "Cal_SNP_EngnRunDt_v__Version_" 
                INNER JOIN "PLN"."SNP_FinGdCmphsvUnPlndDtl_v" "SNP_FinGdCmphsvUnPlndDtl_v" on "Cal_SNP_EngnRunDt_v__Version_"."DivCd" = "SNP_FinGdCmphsvUnPlndDtl_v"."DivCd" and "Cal_SNP_EngnRunDt_v__Version_"."RefSNP_EngnRunDt" = "SNP_FinGdCmphsvUnPlndDtl_v"."RefSNP_EngnRunDt") 
                INNER JOIN "PLN"."LnPln_Rptg_v" "LnPln_Rptg_v" on "SNP_FinGdCmphsvUnPlndDtl_v"."DivCd" = "LnPln_Rptg_v"."GlblDivCd" and "SNP_FinGdCmphsvUnPlndDtl_v"."PlngCtryCd" = "LnPln_Rptg_v"."PlngCtryCd" and "SNP_FinGdCmphsvUnPlndDtl_v"."PlngProdtCd" = "LnPln_Rptg_v"."PlngProdtCd" and "SNP_FinGdCmphsvUnPlndDtl_v"."MatlTypeCd" = "LnPln_Rptg_v"."MatlTypeCd" and "SNP_FinGdCmphsvUnPlndDtl_v"."RgnSeasnMaxSeqNbr" = "LnPln_Rptg_v"."RgnSeasnMaxSeqNbr") 
                INNER JOIN "PLN"."Div_v" "Div_v" on "SNP_FinGdCmphsvUnPlndDtl_v"."DivCd" = "Div_v"."DivCd") 
                LEFT OUTER JOIN "PLN"."Cal_BusFsclN1GregorianCalDt_v" "T3" on "SNP_FinGdCmphsvUnPlndDtl_v"."BusSeasnRlvncDt" = "T3"."BusDt" and 'NF' = "T3"."BusCalCd" and 'Y' = "T3"."FirstDtOfBusSeasnInd") 
                LEFT OUTER JOIN "PLN"."SNP_EnhcdDmndPrty_v" "SNP_EnhcdDmndPrty_v" on "SNP_FinGdCmphsvUnPlndDtl_v"."EnhcdDmndPrtyCd" = "SNP_EnhcdDmndPrty_v"."EnhcdDmndPrtyCd"

        where "SNP_EnhcdDmndPrty_v"."EnhcdDmndPrtyDesc" in ('Standard') and lower ("LnPln_Rptg_v"."GlblMatlIntntDesc") like '%' || lower ('FLYKNIT') || '%' and "LnPln_Rptg_v"."MatlTypeCd" in ('FG   ', 'PROJ ') and "Div_v"."DivCd" in ('20   ') and "SNP_FinGdCmphsvUnPlndDtl_v"."RefSNP_EngnRunDt" = DATE '{max_eng_date}' and "Cal_SNP_EngnRunDt_v__Version_"."RefSNP_EngnRunDt" = DATE '{max_eng_date}'

        group by "Cal_SNP_EngnRunDt_v__Version_"."RefSNP_EngnRunDt", "SNP_EnhcdDmndPrty_v"."EnhcdDmndPrtyDesc", "LnPln_Rptg_v"."MatlTypeCd", "LnPln_Rptg_v"."PO_DevTeamNm", "LnPln_Rptg_v"."GlblMatlIntntDesc", "Div_v"."DivCd", "LnPln_Rptg_v"."StylDsplyCd", "LnPln_Rptg_v"."GlblProdtLglLongNm", "LnPln_Rptg_v"."RgnSeasnSpclOfrgTypeDesc", right("T3"."BusSeasnCd", 2) || left("T3"."BusSeasnCd", 4)

        """

        plugs_query = """

        select 
        "Cal_SNP_EngnRunDt_v__Version_"."RefSNP_EngnRunDt" "Ver_Wk_Dt" , 
        "Div_v"."DivCd" "Prod_Engn_Cd" , 
        "SNP_EnhcdDmndPrty_v"."EnhcdDmndPrtyDesc" "EDP" , 
        "LnPln_Rptg_v"."MO_ExprsLnInd" "MO_Exprs_Ln_Ind" , 
        "LnPln_Rptg_v"."StylDsplyCd" "Styl_Dsply_Cd" , 
        "LnPln_Rptg_v"."ColrDsplyCd" "Colr_Dsply_Cd" , 
        "LnPln_Rptg_v"."PlngProdtCd" "APO_Prod_Cd" , 
        "LnPln_Rptg_v"."GlblProdtLglLongNm" "APO_Prod_Desc" , 
        "LnPln_Rptg_v"."MatlTypeCd" "Matl_Type_Cd" , 
        "LnPln_Rptg_v"."GlblLgDesc" "Lg_Desc" , 
        "LnPln_Rptg_v"."GlblPlngProdtGrpDesc" "Glbl_Plng_Prodt_Grp_Desc" , 
        "LnPln_Rptg_v"."GlblCatCoreFcsDesc" "Gbl_Cat_Core_Focs_Desc" , 
        sum("SNP_FinGdCmphsvUnPlndDtl_v"."TotDmndQty") "Tot_Rlsd_DP_Qty" , 
        sum("SNP_FinGdCmphsvUnPlndDtl_v"."NetRqmntsQty") "Net_Rqrmnts_Qty" , 
        sum("SNP_FinGdCmphsvUnPlndDtl_v"."TotUnPlndDmndQty") "Tot_Unplanned_Dmnd_Qty"

        from (((
            "PLN"."Cal_SNP_EngnRunDt_v" "Cal_SNP_EngnRunDt_v__Version_" 
                INNER JOIN "PLN"."SNP_FinGdCmphsvUnPlndDtl_v" "SNP_FinGdCmphsvUnPlndDtl_v" on "Cal_SNP_EngnRunDt_v__Version_"."DivCd" = "SNP_FinGdCmphsvUnPlndDtl_v"."DivCd" and "Cal_SNP_EngnRunDt_v__Version_"."RefSNP_EngnRunDt" = "SNP_FinGdCmphsvUnPlndDtl_v"."RefSNP_EngnRunDt") 
                INNER JOIN "PLN"."LnPln_Rptg_v" "LnPln_Rptg_v" on "SNP_FinGdCmphsvUnPlndDtl_v"."DivCd" = "LnPln_Rptg_v"."GlblDivCd" and "SNP_FinGdCmphsvUnPlndDtl_v"."PlngCtryCd" = "LnPln_Rptg_v"."PlngCtryCd" and "SNP_FinGdCmphsvUnPlndDtl_v"."PlngProdtCd" = "LnPln_Rptg_v"."PlngProdtCd" and "SNP_FinGdCmphsvUnPlndDtl_v"."MatlTypeCd" = "LnPln_Rptg_v"."MatlTypeCd" and "SNP_FinGdCmphsvUnPlndDtl_v"."RgnSeasnMaxSeqNbr" = "LnPln_Rptg_v"."RgnSeasnMaxSeqNbr") 
                INNER JOIN "PLN"."Div_v" "Div_v" on "SNP_FinGdCmphsvUnPlndDtl_v"."DivCd" = "Div_v"."DivCd") 
                LEFT OUTER JOIN "PLN"."SNP_EnhcdDmndPrty_v" "SNP_EnhcdDmndPrty_v" on "SNP_FinGdCmphsvUnPlndDtl_v"."EnhcdDmndPrtyCd" = "SNP_EnhcdDmndPrty_v"."EnhcdDmndPrtyCd"

        where "SNP_FinGdCmphsvUnPlndDtl_v"."RefSNP_EngnRunDt" = DATE '{max_eng_date}' and "Cal_SNP_EngnRunDt_v__Version_"."RefSNP_EngnRunDt" = DATE '{max_eng_date}' and "LnPln_Rptg_v"."MatlTypeCd" in ('PLUG ') and "SNP_EnhcdDmndPrty_v"."EnhcdDmndPrtyDesc" in ('Standard')
        
        group by "Cal_SNP_EngnRunDt_v__Version_"."RefSNP_EngnRunDt", "SNP_EnhcdDmndPrty_v"."EnhcdDmndPrtyDesc", "LnPln_Rptg_v"."MatlTypeCd", "LnPln_Rptg_v"."MO_ExprsLnInd", "LnPln_Rptg_v"."StylDsplyCd", "Div_v"."DivCd", "LnPln_Rptg_v"."ColrDsplyCd", "LnPln_Rptg_v"."PlngProdtCd", "LnPln_Rptg_v"."GlblProdtLglLongNm", "LnPln_Rptg_v"."GlblLgDesc", "LnPln_Rptg_v"."GlblPlngProdtGrpDesc", "LnPln_Rptg_v"."GlblCatCoreFcsDesc"

        order by 2 asc 

        """

        print('Querying Teradata for Flyknit + Plug styles...')
        df_fly = pd.read_sql_query(flyknit_query.format(max_eng_date = max_engine_run_dt), session)
        df_plug = pd.read_sql_query(plugs_query.format(max_eng_date = max_engine_run_dt), session)
        df_plug['Prod_Engn_Cd'] = df_plug['Prod_Engn_Cd'].astype(int)

        def categorize(row):
            prod_desc = row['APO_Prod_Desc'].lower()
            regexp = re.compile(r'(?:^|\W)promo(?:$|\W)|(?:^|\W)pmo(?:$|\W)|(?:^|\W)gel(?:$|\W)|(?:^|\W)gel\d+(?:$|\W)')
            if regexp.search(prod_desc):
                return 'PLUG | GEL / PROMO'
            else:
                return 'PLUG | OTHER'

        def assign_priority(row):
            check = row['gel/promo check']
            if check == 'PLUG | GEL / PROMO':
                return '50'
            return '150'

        print('Building templates...')
        plug_ap_template = pd.DataFrame(columns=ap_headers)
        plug_fw_template = pd.DataFrame(columns=fw_headers)
        plug_eq_template = pd.DataFrame(columns=fw_headers)

        if len(df_plug) != 0:
            df_plug['gel/promo check'] = df_plug.apply(lambda row: categorize(row), axis=1)
            df_plug['gel/promo priority'] = df_plug.apply(lambda row: assign_priority(row), axis=1)

            df_plug_ap = df_plug[df_plug['Prod_Engn_Cd'] == 10]
            df_plug_fw = df_plug[df_plug['Prod_Engn_Cd'] == 20]
            df_plug_eq = df_plug[df_plug['Prod_Engn_Cd'] == 30]
            
            
        # Apparel plugs
            plug_ap_template['StyleCode'] = df_plug_ap['APO_Prod_Cd']
            plug_ap_template['Reason'] = df_plug_ap['gel/promo check']
            plug_ap_template['PriorityDesc'] = 'P'
            plug_ap_template['Priority'] = df_plug_ap['gel/promo priority']
            plug_ap_template['RequestedBy'] = 'GOVERNANCE STANDARD'
            plug_ap_template['DefaultPriority'] = '100'
            plug_ap_template['updFlag'] = "I"

        # Footwear plugs
            plug_fw_template['StyleCode'] = df_plug_fw['APO_Prod_Cd']
            plug_fw_template['PriorityDesc'] = 'P'
            plug_fw_template['Reason'] = df_plug_fw['gel/promo check']
            plug_fw_template['RequestedBy'] = 'GOVERNANCE STANDARD'
            plug_fw_template['Priority'] = df_plug_fw['gel/promo priority']
            plug_fw_template['DefaultPriority'] = '100'
            plug_fw_template['updFlag'] = "I"

        # # Equipment plugs
            plug_eq_template['StyleCode'] = df_plug_eq['APO_Prod_Cd']
            plug_eq_template['PriorityDesc'] = 'P'
            plug_eq_template['Reason'] = df_plug_eq['gel/promo check']
            plug_eq_template['RequestedBy'] = 'GOVERNANCE STANDARD'
            plug_eq_template['Priority'] = df_plug_eq['gel/promo priority']
            plug_eq_template['DefaultPriority'] = '100'
            plug_eq_template['updFlag'] = "I"

        fly_fw_template = pd.DataFrame(columns=fw_headers)

        if len(df_fly) != 0:
            fly_fw_template['DemandSeason'] = df_fly['SesnYr']
            fly_fw_template['StyleCode'] = df_fly['Styl_Dsply_Cd']
            fly_fw_template['PriorityDesc'] = 'P'
            fly_fw_template['Reason'] = 'FLYKNIT'
            fly_fw_template['RequestedBy'] = 'GOVERNANCE STANDARD'
            fly_fw_template['Priority'] = '50'
            fly_fw_template['DefaultPriority'] = '100'
            fly_fw_template['updFlag'] = "I"

        fw_comb = pd.concat([gel_fw_template, fly_fw_template, plug_fw_template], ignore_index=True)
        ap_comb = pd.concat([gel_ap_template, plug_ap_template], ignore_index=True)
        eq_comb = plug_eq_template

        fw_comb = fw_comb.copy(deep=True)
        ap_comb = ap_comb.copy(deep=True)
        eq_comb = eq_comb.copy(deep=True)

        def write_excel_file(df, pe):
            path = '''C:\\Users\\''' + username + '''\\Box\\Global Supply And Inventory\\Global S&IP Operations\\03 SUPPLY\\OPERATIONS\\PARAMETER MANAGEMENT\\EDP - PRIORITY\\AD HOC\\Combined_Weekly_Upload_Templates\\'''
            
            if pe == 'ap':
                writer = pd.ExcelWriter(path + str(today) + '_Apparel_EDP_Upload' + '.xlsx')
                df.to_excel(writer, index=False)
                writer.close()
            elif pe == 'fw':
                writer = pd.ExcelWriter(path + str(today) + '_Footwear_EDP_Upload' + '.xlsx')
                df.to_excel(writer, index=False)
                writer.close()
            elif pe == 'eq' and len(eq_comb) > 0:
                writer = pd.ExcelWriter(path + str(today) + '_Equipment_EDP_Upload' + '.xlsx')
                df.to_excel(writer, index=False)
                writer.close()

        print('Writing templates to Excel...')

        write_excel_file(ap_comb, 'ap')
        write_excel_file(fw_comb, 'fw')
        write_excel_file(eq_comb, 'eq')

        stop = timeit.default_timer()
        if len(eq_comb) == 0:
            print('\nAP/FW Time: ', stop - start, '\n')
        else:
            print('\nAP/FW + EQ! Time: ', stop - start, '\n')


    except PermissionError:
        print('\nERROR!!!'"\nCannot write to filename that is already open. Please close any open upload templates and rerun the application.\n")
    except BaseException:
        import sys
        print(sys.exc_info()[0])
        import traceback
        print(traceback.format_exc())
    finally:
        print("Press Enter to continue ...")
        input()

