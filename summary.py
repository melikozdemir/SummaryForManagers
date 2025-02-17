import openpyxl
import os
import requests
import json
import re
from mtranslate import translate


folder_path = "C:/Users/Melik/OneDrive/Masaüstü/Günlük Üretim Raporları"

#Klasördeki tüm Excel dosyalarını al
excel_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx")]

#En son değiştirilen Excel dosyasını bul
if not excel_files:
    raise FileNotFoundError("Belirtilen klasörde herhangi bir Excel dosyası bulunamadı.")

latest_file = max(excel_files, key=lambda f: os.path.getmtime(os.path.join(folder_path, f)))
latest_file_path = os.path.join(folder_path, latest_file)

workbook = openpyxl.load_workbook(latest_file_path)
sheet = workbook.active

#Dosya Yapılandırma Kontrolü (Trafik ve Havuz kontrolü)
if sheet["B7"].value != "TRAFİK" or sheet["C7"].value != "HAVUZ":
    print("⚠️ Lütfen belge yapılandırmasını kontrol ediniz.")
    exit()

#Branşlar (Ana Kategoriler -> Trafik-Ziraat)
branch_cells = ["B19", "B25", "B31", "B37", "B43", "B49", "B55", "B61", "B67", "B73"]
branch_values = [sheet[cell].value for cell in branch_cells]

#Havuz Kategorileri
pool_cells = ["C7", "C12"]
pool_values = [sheet[cell].value for cell in pool_cells]

#Alt Kategoriler (Banka, Acente vb.)
sub_branches = ["Banka", "Acenta", "Broker", "Merkez", "Toplam"]

#Sütun Başlıkları
column_headers = ["F5", "G5", "H5", "J5", "K5", "L5", "M5", "N5", "O5", "Q5", "R5", "S5", "T5", "U5", "V5", "W5", "X5"]
cleaned_col_values = [str(sheet[col].value).replace("\n", " ") if isinstance(sheet[col].value, str) else sheet[col].value for col in column_headers]

#Hücre Referanslarını (F19-X77 Arası)
column_letters = [chr(i) for i in range(ord('F'), ord('X') + 1) if chr(i) not in ['I', 'P']]
column_decimals = [i for i in range(19, 78) if i not in [24, 30, 36, 42, 48, 54, 60, 66, 72]]
column_decimals_for_pool = list(range(7, 18))


# Havuz ve Havuz Dışı hücre referansları
cell_references_for_pool = [f"{letter}{decimal}" for decimal in column_decimals_for_pool for letter in column_letters]
cell_references = [f"{letter}{decimal}" for decimal in column_decimals for letter in column_letters]

cell_references2 = cell_references[:]

#Verileri İşleyip Yazdır
print(f"\n **En son güncellenmiş Excel dosyası:** {latest_file} \n")

date = ""
output= ""
date += sheet["B2"].value
date = re.sub(r"[^0-9.]", "", date)
date+= "\n"


cnt = 0

#Havuz Verileri Yazdır
for pool in pool_values:
    for sub_branch in sub_branches:
        for col_name, cell_ref in zip(cleaned_col_values, cell_references_for_pool):
            value = sheet[cell_ref].value
            output += f"Trafik branşının {pool} kısmında {sub_branch} dalındaki {col_name} üretim değeri: {value} \n"
            cnt += 1
            if cnt == 17:
                del cell_references_for_pool[:17]
                cnt = 0
            #print(cell_ref)

ccnt = 0

#Havuz ve Havuz Dışı Toplam Değerleri Yazdır
for col_name, cell_ref in zip(cleaned_col_values, cell_references2):
    value = sheet[cell_ref].value
    output += f"Trafik dalının Havuz ve Havuz dışındaki Toplam {col_name} üretim değeri: {value} \n"
    ccnt+=1
    if ccnt == 17:
        del cell_references2[:17]
        ccnt = 0

#print(cell_references)
counter = 0

#Her Branş İçin Üretim Verilerini Yazdır
for branch in branch_values:
    for sub_branch in sub_branches:
        for col_name, cell_ref in zip(cleaned_col_values, cell_references):
            value = sheet[cell_ref].value
            output += f"{branch} branşının {sub_branch} dalındaki {col_name} üretim değeri: {value} \n"
            counter += 1
            if counter == 17:
                del cell_references[:17]
                counter = 0
            #print(cell_ref)


output = output.replace(".", ",")

with open("data.txt", "w", encoding="utf-8") as file:
    file.write(date + output)

output = "\n".join(output.splitlines()[:187])


lm_studio_url = "http://localhost:1234/v1/chat/completions"  # LM Studio API adresi

headers = {
    "Content-Type": "application/json"

}


lines = output.strip().split("\n")  # Satır bazlı böl
chunk_size = 17
original_text = ""
translated_text = ""

try:
    for i in range(0, len(lines), chunk_size):
        chunk = lines[i:i + chunk_size]  # 17'şer satırlık parçalar oluştur
        formatted_text = "\n".join(chunk)  # 17 satırlık parçayı tekrar birleştir
        print(formatted_text)
        eng_text = translate(formatted_text, "en")
        payload = {
    "model": "openhermes-2.5-mistral-7b-16k",
    "messages": [
      { "role": "system", "content": "Sen bir Sigorta şirketi yöneticileri için üretim raporlarından özet hazırlayan bir asistansın. Aşağıdaki üretim verilerini yalnızca en önemli noktaları vurgulayarak özetle. Gereksiz detayları çıkar ve yalnızca yöneticinin karar vermesi için kritik bilgileri sun. Aşağıdaki metni 4 cümle ile özetle." },
      { "role": "user", "content":  f" Lütfen bu üretim verilerinden bir yönetici özeti hazırla. \n {formatted_text}" }
    ],

    "max_tokens": 300,
    "temperature": 0,
    "top_p": 0.9,
    "top_k": 20
    
}
  # API için uygun payload
        response = requests.post(lm_studio_url, headers=headers, data=json.dumps(payload))

        response_data = response.json()

        if "choices" in response_data and len(response_data["choices"]) > 0:
            response_text = response_data["choices"][0].get("text", "")
            original_text += response_text + "\n"
            translated_text += translate(response_text, "tr")

        else:
            print("Beklenen formatta yanıt alınamadı:", response_data)

    # Yanıtı yazdır
    with open("Turkish_summary.txt", "w", encoding="utf-8") as file:
        file.write(translated_text + "\n")

    print("\n LM Studio'dan Gelen Yanıt:")
    print(original_text)

except Exception as e:
    print("API isteği başarısız oldu:", e)
