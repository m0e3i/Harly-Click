"""يعمل هاش لكلمة سر الإدارة عشان تحطه في .env (من غير ما تكتب الباسورد نفسه في أي ملف)
الاستخدام:  python make_admin_hash.py
"""
import getpass
from werkzeug.security import generate_password_hash

pw = getpass.getpass("اكتب كلمة سر الإدارة الجديدة: ")
if len(pw) < 8:
    raise SystemExit("لازم 8 حروف على الأقل")
if pw != getpass.getpass("اكتبها تاني للتأكيد: "):
    raise SystemExit("مش متطابقتين")
print("\nانسخ السطر ده في ملف .env (بدل سطر ADMIN_PASS_HASH):\n")
print("ADMIN_PASS_HASH=" + generate_password_hash(pw))
