from flask import Flask, render_template, jsonify, request
import os
import shutil
import pythoncom
import win32com.client
from werkzeug.utils import secure_filename

app = Flask(__name__)

BASE_DIR = 'static'
PATHS = {
    'videos': os.path.join(BASE_DIR, 'videos'),
    'images': os.path.join(BASE_DIR, 'images'),
    'audios': os.path.join(BASE_DIR, 'audios')
}

@app.route('/panel')
def index():
    return render_template('panel.html')

@app.route('/ekran')
def ekran():
    return render_template('ekran.html')

@app.route('/api/media')
def get_media():
    data = {}
    for key, path in PATHS.items():
        if not os.path.exists(path): os.makedirs(path)
        data[key] = [f for f in os.listdir(path) if not f.startswith('.')]
    return jsonify(data)

@app.route('/api/upload_pptx', methods=['POST'])
def upload_pptx():
    if 'file' not in request.files:
        return jsonify({'error': 'Dosya bulunamadı'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'Dosya seçilmedi'}), 400
    
    if file and file.filename.endswith('.pptx'):
        # Dosya adını Türkçe karakter ve boşluklardan arındırarak güvenli hale getir
        filename = secure_filename(file.filename)
        temp_path = os.path.abspath(os.path.join(BASE_DIR, filename))
        file.save(temp_path)
        
        # Sunumun adı (Uzantısız)
        pres_name = filename.replace('.pptx', '')
        # PowerPoint'in dışa aktaracağı geçici klasör
        temp_out_folder = os.path.abspath(os.path.join(PATHS['images'], pres_name))
        
        try:
            # Arka planda PowerPoint'i başlat
            pythoncom.CoInitialize()
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            
            # Sunumu penceresiz olarak aç
            presentation = powerpoint.Presentations.Open(temp_path, WithWindow=False)
            
            # 17 numaralı format JPEG'dir. Sunumu JPEG olarak dışa aktar.
            presentation.SaveAs(temp_out_folder, 17)
            presentation.Close()
            powerpoint.Quit()
            pythoncom.CoUninitialize()
            
            # PPTX'in oluşturduğu klasördeki resimleri ana 'images' klasörüne taşı ve yeniden adlandır
            for slide_img in os.listdir(temp_out_folder):
                old_img_path = os.path.join(temp_out_folder, slide_img)
                # Resim ismini "SunumAdi_Slayt1.JPG" formatına çevir
                new_img_name = f"{pres_name}_{slide_img}" 
                new_img_path = os.path.join(PATHS['images'], new_img_name)
                shutil.move(old_img_path, new_img_path)
            
            # Boşalan geçici klasörü ve ilk yüklenen pptx dosyasını sil
            os.rmdir(temp_out_folder)
            os.remove(temp_path)
            
            return jsonify({'success': True, 'message': 'Dönüştürme başarılı!'})
            
        except Exception as e:
            pythoncom.CoUninitialize()
            return jsonify({'error': str(e)}), 500
    else:
        return jsonify({'error': 'Sadece .pptx uzantılı dosyalar desteklenir'}), 400

if __name__ == '__main__':
    app.run(debug=True, port=5000)