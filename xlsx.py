import pandas as pd

class text :
    contact = ""
    status = ""
    type = ""
    contact_file_name = ""
    status_file_name = ""
    type_file_name = ""
    postal_filename = ""

    def __init__(self,contact_path=None,status_path=None,type_path=None,postal_filename =None) -> None:
        self.contact_file_name = contact_path
        self.status_file_name = status_path
        self.type_file_name = type_path
        self.postal_filename = postal_filename

    def read_status(self):
        try:
            with open(self.status_file_name, 'r', encoding='utf-8') as file:
                # ใช้คำสั่ง read() เพื่ออ่านข้อมูลจากไฟล์
                data = file.read()
                # print(data)
                return data
        except FileNotFoundError:
            # print(f'ไม่พบไฟล์ที่ชื่อ {self.contact_file_name}')
            pass
        except Exception as e:
            # print(f'เกิดข้อผิดพลาด: {e}')
            pass

    def read_type(self):
        try:
            with open(self.type_file_name, 'r', encoding='utf-8') as file:
                # ใช้คำสั่ง read() เพื่ออ่านข้อมูลจากไฟล์
                data = file.read()
                # print(data)
                return data
        except FileNotFoundError:
            # print(f'ไม่พบไฟล์ที่ชื่อ {self.contact_file_name}')
            pass
        except Exception as e:
            # print(f'เกิดข้อผิดพลาด: {e}')
            pass

    def read_contact(self):
        try:
            with open(self.contact_file_name, 'r', encoding='utf-8') as file:
                # ใช้คำสั่ง read() เพื่ออ่านข้อมูลจากไฟล์
                data = file.read()
                # print(data)
                return data
        except FileNotFoundError:
            # print(f'ไม่พบไฟล์ที่ชื่อ {self.contact_file_name}')
            pass
        except Exception as e:
            # print(f'เกิดข้อผิดพลาด: {e}')
            pass
    def read_postal(self):
        try:
            with open(self.postal_filename, 'r', encoding='utf-8') as file:
                # ใช้คำสั่ง read() เพื่ออ่านข้อมูลจากไฟล์
                data = file.read()
                # print(data)
                return data
        except FileNotFoundError:
            # print(f'ไม่พบไฟล์ที่ชื่อ {self.contact_file_name}')
            pass
        except Exception as e:
            # print(f'เกิดข้อผิดพลาด: {e}')
            pass


class xlsx :
    Text  = text('contact.txt','status.txt','type.txt','postal_code.txt')
    fileName =""
    result_dict_edit = {"code":"",
                   "condotel":"",
                   "floor":"",
                   "building":"",
                   "type":"",
                   "size":"",
                   "price":"",
                   "key":"",
                   "sales":"",
                   "rent":"",
                   "status":"",
                   "fq":"",
                   "tq":"",
                   "company_name":"",
                   "remark":"",
                   "bed":"",
                   "bath":"",
                   "contact_name":"",
                   "contact_email":"",
                   "contact_tel":"",
                   "contact_info":"",
                   "type_text":"",
                   "postal_code":""
                   }
    
    def __init__(self,file_name) -> None:
        self.fileName = file_name

    def read(self):
        self.data = pd.read_excel(self.fileName)
        return self.data
    
    def search(self,row,columns):
        result = self.data.iloc[row,columns]
        return result
    
    def serach_row(self,row):
        result = self.data.iloc[row, :]
        return result
    
    def search_columns(self,columns):
        result = self.data.iloc[:,columns]
        return result
    
    # def get_value_row(self,row):
    #     data = self.data.iloc[row, :]
    #     for index, value in data.items():
    #         print(f"{index}: {value}")

    def get_value_row(self, row):
        data = self.data.iloc[row, :]     
        # สร้าง Dictionary เปล่าขึ้นมา
        result_dict = {}
        # วนลูปผ่านคอลัมน์และค่าในแต่ละคอลัมน์ของ DataFrame
        for index, value in data.items():
            # เพิ่มคีย์และค่าลงใน Dictionary
            result_dict[index] = value  
        # print(result_dict['Building '])
        # print("-------------------------------------------------------------------------------------")
        data_bed_bath = result_dict['Type']
        data_bed_bath = data_bed_bath.split()
        #print(data_bed_bath)
        self.result_dict_edit['code'] = result_dict['code']
        self.result_dict_edit['condotel'] = result_dict['condotel']
        self.result_dict_edit['floor'] = result_dict['Floor']
        self.result_dict_edit['building'] = result_dict['Building ']
        self.result_dict_edit['type'] = result_dict['Type']
        self.result_dict_edit['size'] = result_dict['Size']
        self.result_dict_edit['price'] =result_dict['Price']
        self.result_dict_edit['key'] = result_dict['Key']
        self.result_dict_edit['sales'] = result_dict['Sales']
        self.result_dict_edit['rent'] = result_dict['Rent']
        self.result_dict_edit['status'] = self.Text.read_status()
        self.result_dict_edit['fq'] = result_dict['FQ']
        self.result_dict_edit['tq'] =result_dict['TQ']
        self.result_dict_edit['company_name'] = result_dict['Company Name']
        self.result_dict_edit['remark'] = result_dict['Remark']
        self.result_dict_edit['bed'] = data_bed_bath[0]
        self.result_dict_edit['bath'] = data_bed_bath[2]
        data_contact = self.read_txt_contact()
        self.result_dict_edit['contact_name'] = data_contact['name']
        self.result_dict_edit['contact_email'] = data_contact['email']
        self.result_dict_edit['contact_tel'] = data_contact['tel']
        self.result_dict_edit['contact_info'] = data_contact['info']
        self.result_dict_edit['type_text'] = self.Text.read_type()
        self.result_dict_edit['postal_code'] = self.Text.read_postal()
        # print("bed: "+self.result_dict_edit['bed'])
        # print("bath: "+self.result_dict_edit['bath'])
        return self.result_dict_edit
    
    def read_txt_contact(self):
        dict_result = {"name":"","email":"","tel":"","info":""}
        name=""
        email=""
        tel=""
        info=""

        result = self.Text.read_contact()
        result = result.split(",")

        name = result[0]
        email= result[1]
        tel = result[2]
        info = result[3]

        name = name.split("=")
        email = email.split("=")
        tel = tel.split("=")
        info = info.split("=")

        dict_result['name'] = name[1]
        dict_result['email'] = email[1]
        dict_result['tel'] = tel[1]
        dict_result['info'] = info[1]

        #print(name,email,tel,info)
        return dict_result
    
    def read_txt_status(self):
        result = self.Text.read_status()
        return result
    
    def read_txt_type(self):
        result = self.Text.read_type()
        return result 
        
    def read_txt_postal(self):

        result = self.Text.read_postal()
        return result


