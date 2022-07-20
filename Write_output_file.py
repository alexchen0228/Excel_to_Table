class write_output_file():
    def __init__(self,sheet_obj):
        self.table_name    = sheet_obj['B4'].value
        self.table_comment = sheet_obj['B5'].value