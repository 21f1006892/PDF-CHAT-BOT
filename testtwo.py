import pandas as pd

# ===== FILE PATHS =====
SOURCE_FILE = "SEBI_Monthly_Portfolio 31 JAN 2026.xls"
MASTER_FILE = "IN_MF_PORTFOLIO_DETAILS_ACROSS_SCHEMES.xlsx"

scheme_dict = {'ABBSEIIF': 'EQUITY', 'ABSLBCF': 'EQUITY', 'ABSLCONF': 'EQUITY', 'ABSLESG': 'EQUITY', 
'ABSLLDF': 'DEBT', 'ABSLMAAF': 'HYBRID', 'ABSLMCF': 'EQUITY', 'ABSLQF': 'EQUITY', 'ABSLSO': 'EQUITY', 
'ABSLTNLF': 'EQUITY', 'ABSLUS03': 'FOF', 'ABSLUS10': 'FOF', 'ADVG': 'EQUITY', 'BANKETF': 'EQUITY', 'BBIF': 'DEBT', 
'BBP': 'DEBT', 'BDB': 'DEBT', 'BDYP': 'EQUITY', 'BFL': 'DEBT', 'BFS': 'DEBT', 'BINFRA': 'EQUITY', 
'BINTEQA': 'INTRNL', 'BMIDX50': 'EQUITY', 'BQIDX50': 'EQUITY', 'BSL95F': 'HYBRID', 'BSLAAMM': 'FOF', 
'BSLADMM': 'FOF', 'BSLBBYW': 'EQUITY', 'BSLBKFS': 'EQUITY', 'BSLCBF': 'DEBT', 'BSLCM': 'DEBT', 'BSLDAAF': 'HYBRID', 
'BSLEAF': 'HYBRID', 'BSLEQSF': 'HYBRID', 'BSLEQTY': 'EQUITY', 'BSLFEF': 'EQUITY', 'BSLFPAP': 'FOF', 'BSLFPCP': 'FOF',
 'BSLFPPP': 'FOF', 'BSLIF': 'DEBT', 'BSLMFG': 'EQUITY', 'BSLMIFOF': 'FOF', 'BSLMTP': 'DEBT', 'BSLNMF': 'EQUITY', 
 'BSLONF': 'DEBT', 'BSLPHF': 'EQUITY', 'BSLR96': 'EQUITY', 'BSLRF30': 'HYBRID', 'BSLRF40': 'HYBRID', 
 'BSLRF50': 'HYBRID', 'BSLRF50P': 'DEBT', 'BSLSTF': 'DEBT', 'BSLTA1': 'EQUITY', 'BTOP100': 'EQUITY', 
 'C10YGETF': 'DEBT', 'CASH': 'DEBT', 'CBGETF': 'DEBT', 'CIGAPR26': 'DEBT', 'CIGAPR28': 'DEBT', 'CIGAPR29': 'DEBT', 
 'CIGAPR33': 'DEBT', 'CIGJUN27': 'DEBT', 'CIGPSA28': 'DEBT', 'CISJUN32': 'DEBT', 'COFAIETF': 'DEBT', 
 'CSFSD12': 'DEBT', 'CSFSD6': 'DEBT', 'CSFSI27': 'DEBT', 'CSNHFS26': 'DEBT', 'CSPAPA26': 'DEBT', 'CSPAPA27': 'DEBT', 
 'FTPTI': 'DEBT', 'FTPTJ': 'DEBT', 'FTPTQ': 'DEBT', 'FTPUB': 'DEBT', 'FTPUJ': 'DEBT', 'GENNEXT': 'EQUITY', 
 'GOLDETF': 'GOLD', 'GOLDFOF': 'FOF', 'INV': 'DEBT', 'MIDCAP': 'EQUITY', 'MIP25': 'HYBRID', 'MNC': 'EQUITY', 
 'N30MTETF': 'EQUITY', 'N30QTETF': 'EQUITY', 'NEWIF50': 'EQUITY', 'NHLTHETF': 'EQUITY', 'NIFTY': 'EQUITY', 
 'NIFTYETF': 'EQUITY', 'NIFYNX50': 'EQUITY', 'NINDDEF': 'EQUITY', 'NITETF': 'EQUITY', 'NMID150': 'EQUITY', 
 'NPSEETF': 'EQUITY', 'NSDAPR27': 'DEBT', 'NSDAQ100': 'FOF', 'NSDSEP27': 'DEBT', 'NSMALL50': 'EQUITY', 
 'NSPPBS26': 'DEBT', 'NXTIDX50': 'EQUITY', 'PLUS': 'DEBT', 'PSUEQ': 'EQUITY', 'PURE': 'EQUITY', 'SENSXETF': 'EQUITY',
  'SILVRETF': 'SILVER', 'SILVRFOF': 'FOF', 'BSLGCF': 'OTH', 'BSLGRE': 'OTH'}


source_file = pd.ExcelFile(SOURCE_FILE, engine='xlrd')
master_file = pd.read_excel(MASTER_FILE)

# print(master_file.columns)
for sheet_name in source_file.sheet_names:
    if sheet_name=='Index':
        continue
    if scheme_dict[sheet_name] == 'EQUITY' or scheme_dict[sheet_name] == 'INTRNL':
        data = pd.read_excel(SOURCE_FILE, sheet_name=sheet_name, header=None, usecols="B:G", engine='xlrd')
        
        # Find the row index containing "Name of the Instrument" (header row)
        header_row_index = data[data.apply(lambda row: row.astype(str).str.contains("Name of the Instrument", case=False).any(), axis=1)].index[0]
        
        # Find the row index containing "GRAND TOTAL" (last row of table)
        last_row_index = data[data.apply(lambda row: row.astype(str).str.contains("GRAND TOTAL", case=False).any(), axis=1)].index[0]
        
        # Slice the dataframe: from header row to last row
        df_table = data.iloc[header_row_index:last_row_index + 1].reset_index(drop=True)

        # Set the header row
        df_table.columns = df_table.iloc[0]
        df_table = df_table[1:]  # Remove the header row from data
        
        # Reset index
        df_table = df_table.reset_index(drop=True)

        # Search string (partial match)
        start_search_string = "(a) Listed / awaiting listing on Stock"
        end_search_string = "Sub Total"

        # Find the row index in the first column where the string occurs
        start_marker_row_index = df_table[df_table.iloc[:, 0].astype(str).str.contains(start_search_string, case=False, na=False, regex=False)].index
        end_marker_row_index = df_table[df_table.iloc[:, 0].astype(str).str.contains(end_search_string, case=False, na=False, regex=False)].index
        start_marker_row_index = start_marker_row_index[0]  # take the first match
        end_marker_row_index = end_marker_row_index[0]  # take the first match
        source_table = df_table.loc[(start_marker_row_index+1):(end_marker_row_index-1)].reset_index(drop=True) 
        source_table.columns.name = None

        # Read data from master file for the ongoing scheme
        master_data = pd.read_excel(MASTER_FILE, usecols=["Client Code","Issuer Name","Security Type Name","ISIN","Industry","Quantity","Total Market Value (Rs.)","% to Net assests"])
        if scheme_dict[sheet_name]=='INTRNL':
            master_table = master_data[(master_data["Client Code"]==sheet_name)&(master_data["Security Type Name"].isin(["International Equity"]))]
        if scheme_dict[sheet_name]=='EQUITY':    
            master_table = master_data[(master_data["Client Code"]==sheet_name)&(master_data["Security Type Name"].isin(["EQUITY", "PREFERRED STOCK"]))]
        master_table = master_table.drop(columns=["Client Code","Security Type Name"])
        master_table = master_table.rename(columns={"Issuer Name":"Name of the Instrument", "Industry":"Industry^ / Rating", "Total Market Value (Rs.)":"Market/Fair Value\r\n(Rs.in Lacs)", "% to Net assests":"% to Net Assets"})[['Name of the Instrument', 'ISIN', 'Industry^ / Rating', 'Quantity',
       'Market/Fair Value\r\n(Rs.in Lacs)', '% to Net Assets']]
        master_table = master_table.reset_index(drop=True)
        master_table['Market/Fair Value\r\n(Rs.in Lacs)'] = (master_table['Market/Fair Value\r\n(Rs.in Lacs)'] / 100000).round(2)
        # master_table['% to Net Assets'] = (master_table['% to Net Assets'] / 100)
        # print(master_table.shape)
        # print(source_table.shape)
        # print(source_table.merge(master_table))
        df1, df2 = master_table.align(source_table)
        diff = df1.ne(df2)
        mismatch_cols = diff.apply(lambda row: list(df1.columns[row]), axis=1)
        comparison = master_table.merge(source_table, on="ISIN", how='outer', indicator=True, suffixes=("_master", "_source"))
        cols_to_check = [
            "Name of the Instrument",
            "Industry^ / Rating",
            "Quantity",
            "Market/Fair Value\r\n(Rs.in Lacs)",
            "% to Net Assets"
        ]
        def get_mismatched_columns(row):
            mismatches = []
            
            for col in cols_to_check:
                master_col = col + "_master"
                source_col = col + "_source"
                
                if row.get(master_col) != row.get(source_col):
                    mismatches.append(col)
            
            return mismatches
        comparison["Mismatched Columns"] = comparison.apply(get_mismatched_columns, axis=1)
        result = comparison[comparison["_merge"] != "both"]
        if (result.shape[0])>0:
            result.to_excel(f"output/mismatched_rows_{sheet_name}.xlsx", index=False)
        
        print(source_table.tail())
        print(master_table.tail())
        break

        # n = 1
        # while df_table.iloc[marker_row_index+n]['Name of the Instrument']!='Sub Total':
        #     print(master_table['ISIN'].isin([df_table.iloc[marker_row_index+n]['ISIN']]).any())

        #     n+=1