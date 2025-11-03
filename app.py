import os
import json
from datetime import datetime
import csv
from pathlib import Path
from flask import Flask, request, jsonify, send_file, render_template
from flask_cors import CORS
from werkzeug.utils import secure_filename
import httpx
from openai import OpenAI
from pdf_abstract import process_path

app = Flask(__name__)
CORS(app)

# Setup upload folder
UPLOAD_FOLDER = 'pdf_uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def generate_keywords_with_llm(title, author, year, abstract, api_key):
    """Uses an LLM to generate keywords from paper metadata."""
    if not all([title, abstract, api_key]):
        return "（資料不齊全，略過 AI 分析）"
    
    try:
        prompt = f"""請根據以下論文資訊，分析並提取核心主題與目標的關鍵字。
        
論文資訊:
題目: {title}
作者: {author}
年份: {year}
摘要: {abstract}

請以最簡潔的方式列出這篇論文的核心主題與目標，並僅使用中文頓號 (、) 分隔關鍵字。
格式範例: \"主題1、主題2、主題3\"

請直接回答關鍵字，不要包含任何其他說明或前綴文字。"""

        client = OpenAI(
            base_url="https://models.inference.ai.azure.com",
            api_key=api_key,
            http_client=httpx.Client(proxies={})
        )
        
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": "你是一位專業的學術研究助理，擅長從論文資訊中精準提取核心主題與研究目標。"},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7,
            max_tokens=300
        )
        
        return response.choices[0].message.content.strip()
        
    except Exception as e:
        print(f"LLM API Call Error: {e}")
        return f"（AI 分析失敗：{str(e)}）"

@app.route('/')
def index():
    """Serves the main page."""
    return render_template('index.html')

@app.route('/api/process_batch', methods=['POST'])
def process_batch():
    """Handles the entire batch processing workflow."""
    try:
        files = request.files.getlist('pdfs')
        api_key = request.form.get('apiKey')

        if not files or files[0].filename == '':
            return jsonify({'success': False, 'error': '沒有選擇任何檔案'}), 400
        if not api_key:
            return jsonify({'success': False, 'error': '缺少 API Key'}), 400

        all_records = []
        for file in files:
            if file:
                filename = file.filename
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(filepath)

                pdf_path = Path(filepath)
                pdf_path = Path(filepath)
                success, title, year, author , abstract = process_path(pdf_path,output_dir=Path('abstract_output'), recursive=False, verbose=False, very_verbose=False,to_csv=False)
            
                keywords = generate_keywords_with_llm(title, author, year, abstract, api_key)

                all_records.append({
                    'title': title or '（標題讀取失敗）',
                    'author': author or '（作者讀取失敗）',
                    'year': year or '（年份讀取失敗）',
                    'abstract': abstract or '（摘要讀取失敗）',
                    'keywords': keywords
                })
        
        return jsonify({'success': True, 'records': all_records})

    except Exception as e:
        print(f"Batch Processing Error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/export', methods=['POST'])
def export_csv():
    """Exports records to a CSV file."""
    try:
        data = request.json
        if 'records' not in data or not data['records']:
            return jsonify({'success': False, 'error': '沒有資料可以匯出'}), 400
        
        output_dir = 'output'
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f'論文資料標籤_輸出_{timestamp}.csv'
        filepath = os.path.join(output_dir, filename)
        
        with open(filepath, 'w', newline='', encoding='utf-8-sig') as csvfile:
            fieldnames = ['題目', '作者', '年份', '摘要', '核心主題與目標']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            
            writer.writeheader()
            for record in data['records']:
                writer.writerow({
                    '題目': record.get('title'),
                    '作者': record.get('author'),
                    '年份': record.get('year'),
                    '摘要': record.get('abstract'),
                    '核心主題與目標': record.get('keywords')
                })
        
        return send_file(
            filepath,
            mimetype='text/csv',
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500
if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)