from db_handler import db_handler

db = db_handler('specs.db3')

#db.attach_to_spec("Q:\\Q10. QC-2023\\Q11. THAM DINH TC-Thanh Ha\\1. THẨM ĐỊNH TIÊU CHUẨN-HƯNG YÊN\\Emflotra 10 mg\\1. Đăng ký mới\\5. Emflotra VBF-Gui NC&HY 22.06.23\\1. TCCS- ĐỀ CƯƠNG- BÁO CÁO\\1. Tiêu chuẩn Emflotra 10 mg.txt", 
#	('EMFLOTRA10', '04-ĐS-173-23THY', 3), 'txt_file')
db.compare_spec(('EMFLOTRA25', 'V1/23.12.22', 1), ('EMFLOTRA25', '06.03.24.V2', 1))
# for i in [('EMFLOTRA10', 'V0/29.10.22', 1), 
# 		  ('EMFLOTRA10', 'V1/23.12.22', 1), 
# 		  ('EMFLOTRA10', '06.03.24.V2', 1), 
# 		  ('EMFLOTRA25', 'V0/29.10.22', 1), 
# 		  ('EMFLOTRA25', 'V1/23.12.22', 1), 
# 		  ('EMFLOTRA25', '06.03.24.V2', 1)]:
# 	db.spec_doc_to_txt(i)