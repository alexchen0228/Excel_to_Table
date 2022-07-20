from openpyxl import load_workbook
import os 
import re 


def list_excel_files(dirpath):
    list_all_files = os.listdir(dirpath)
    excel_files = []
    for file in list_all_files:
        if '.xlsx' in file:
            excel_files.append(file)
    return excel_files

def excel_sheetnames_locate(excel_obj):
    book_sheet_names = excel_obj.sheetnames
    excel_file_in_list = []
    regex = re.compile(r"附件1-\d\Z")
    for sheet in book_sheet_names:
        match = regex.search(sheet)
        if match is not None:
        #    print(match.group(0))
            excel_file_in_list.append(match.group(0))
    return excel_file_in_list

class write_output_file():
    def __init__(self,sheet_obj):
        self.sheet_obj           = sheet_obj
        self.table_name          = self.sheet_obj['B4'].value
        self.table_comment       = self.sheet_obj['B5'].value
        self.table_schema        = self.sheet_obj['B1'].value
        self.Table_Space_Value   = self.sheet_obj['G1'].value

    def output_script_part1(self,input_author,input_Issue):
        output_script_part1 = """--===================================================================\n--AUTHOR      : {author}\n--COMMENT     : {Issue} {table_name}  {comment}\n--===================================================================\nCREATE  TABLE WMSR6USR.{table_name} (\n """.format(author = input_author,Issue = input_Issue,table_name =self.table_name,comment =self.table_comment)
        return output_script_part1
    
    def set_col_value(self,col_pre,col_remove_Add):
            col_remove = '{col_pre}{col_remove}'.format(col_pre = col_pre,col_remove = col_remove_Add)
            col_value  = self.sheet_obj[col_remove].value
            return col_value

    def output_script_part2(self):
        col_Name_Value        = self.sheet_obj['A7'].value 
        col_Type_Value        = self.sheet_obj['B7'].value
        col_Null_Value        = self.sheet_obj['E7'].value
        col_Default_Value     = self.sheet_obj['F7'].value 
        col_Pk_Value          = self.sheet_obj['C7'].value
        col_Remark_Value      = self.sheet_obj['G7'].value
        col_all_list          = []
        col_Remark_Value_list = []
        output_script_part2   = []
        Pk_col_name_value     = []
        col_locate = 7


        while col_Name_Value:
            col_all_list.append([col_Name_Value,col_Type_Value,col_Null_Value,col_Default_Value])
            col_Remark_Value_list.append([col_Name_Value,col_Remark_Value])

            col_Name_Value    = self.set_col_value('A',col_remove_Add=col_locate)
            col_Type_Value    = self.set_col_value('B',col_remove_Add=col_locate)
            col_Null_Value    = self.set_col_value('E',col_remove_Add=col_locate)
            col_Default_Value = self.set_col_value('F',col_remove_Add=col_locate)
            col_Pk_Value      = self.set_col_value('C',col_remove_Add=col_locate)
            col_Remark_Value  = self.set_col_value('G',col_remove_Add=col_locate)

            if col_Pk_Value:
                Pk_col_name_value.append(col_Name_Value)
            col_locate +=1

        for row in col_all_list:
            col_value = (' 	{col_name} {col_type} {Set_null} {Set_default} {last_word}\n '.format(col_name = row[0],col_type = row[1],Set_null = '' if row[2] == None else row[2] 
            ,Set_default = '' if row[3] == None else ('DEFAULT {Default}').format(Default = row[3]),last_word = '' if row == col_all_list[-1] else ','  ))
            output_script_part2.append(col_value)
        
        col_index_value   = """	) COMPRESS YES ADAPTIVE IN {Table_space} INDEX IN TS_WMS_BAT_IDX ORGANIZE BY ROW  @\n\n""".format(Table_space =  self.Table_Space_Value )
        output_script_part2.append(col_index_value)

        Pk_col_name_value = map(( lambda x: '\"' + x + '\"'), Pk_col_name_value)
        if Pk_col_name_value:
            col_Pk_Value = """ALTER TABLE {Table_schema}.{Table_name} ADD CONSTRANT PK_{Table_name} PRIMARY KEY({Pk_col_name_value}) @\n""".format(Table_schema = self.table_schema,Table_name =self.table_name,Pk_col_name_value = ",".join(Pk_col_name_value))
            output_script_part2.append(col_Pk_Value)

        col_Remark_table_value = """COMMENT ON TABLE \"{Table_schema}\".\"{Table_name}\" IS\'{Table_comment}\'@\n""".format(Table_schema =self.table_schema,Table_name = self.table_name,Table_comment = self.table_comment)
        output_script_part2.append(col_Remark_table_value)

        for col_remark in col_Remark_Value_list:
            col_remark_value = """COMMENT ON COLUMN \"{Table_schema}\".\"{Table_name}\".\"{Table_col}\" IS \'{Table_remark}\'@\n""".format(Table_schema =self.table_schema,Table_name = self.table_name,Table_col = col_remark[0],Table_remark = col_remark[1])
            output_script_part2.append(col_remark_value)

        return " ".join(output_script_part2)
    
    def output_script_part3(self):
        col_grant_locate = 28
        grant_permission_user = self.sheet_obj['A28'].value
        grant_verb            = self.sheet_obj['B28'].value
        grant_user_list       = []
        output_script_part3   = []

        while grant_permission_user:
            grant_user_list.append([grant_verb,grant_permission_user])
            grant_permission_user    = self.set_col_value('A',col_remove_Add=col_grant_locate)
            grant_verb               = self.set_col_value('B',col_remove_Add=col_grant_locate)
            col_grant_locate +=1
        
        for row in grant_user_list:
            grant_value = """GRANT {grant_verb} ON TABLE \"{table_schema}\".\"{table_name}\" TO USER \"{grant_permission_user}\" @\n""".format(grant_verb = row[0],table_schema =self.table_schema,table_name = self.table_name, grant_permission_user = row[1])
            output_script_part3.append(grant_value)
        return " ".join(output_script_part3)


    def output_write_txt_file(self,output_scirpt_part1,output_script_part2,output_script_part3):
        file_name = '{table_name}.Table.sql'.format(table_name = self.table_name)
        with open(file=file_name,mode='w+',encoding='utf-8') as f :
            f.write(output_scirpt_part1)
            f.write(output_script_part2)
            f.write(output_script_part3)


    
if __name__ == '__main__':
    input_dir_path = input('告訴我，你要create TABLE的路徑?\n')
    input_author   = input('請告訴你是誰?\n')
    input_Issue    = input('請告訴我單號，不然誰知道是哪張單阿?\n')
    excel_files = list_excel_files(input_dir_path)
    for excel_file in excel_files:
        book = load_workbook(excel_file)
        print('正在處理的Excel文件:{excel_file}'.format(excel_file = excel_file))
        sheet_list = excel_sheetnames_locate(excel_obj = book)
        for sheet in sheet_list :
            write_output_file_obj = write_output_file(sheet_obj=book[sheet])
            write_output_file_obj.output_write_txt_file(output_scirpt_part1=write_output_file_obj.output_script_part1(input_author=input_author,input_Issue=input_Issue),output_script_part2=write_output_file_obj.output_script_part2(),output_script_part3=write_output_file_obj.output_script_part3())
