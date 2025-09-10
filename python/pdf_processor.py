#!/usr/bin/env python3
"""
PDF分割・リネームツール
dlフォルダ内のZIPファイルを処理して、outputフォルダに結果を出力
"""

import os
import sys
import zipfile
import tempfile
import shutil
from pathlib import Path
from typing import List, Dict, Tuple
import re
import PyPDF2
from datetime import datetime

class PDFProcessor:
    def __init__(self, dl_dir: str = "../dl", output_dir: str = "../output"):
        self.dl_dir = Path(dl_dir).resolve()
        self.output_dir = Path(output_dir).resolve()
        self.temp_dir = None
        
        # 出力ディレクトリを作成
        self.output_dir.mkdir(exist_ok=True)
        
        print(f"📁 入力ディレクトリ: {self.dl_dir}")
        print(f"📁 出力ディレクトリ: {self.output_dir}")
    
    def get_processing_year_month(self) -> Tuple[int, int]:
        """実行日の1月前の年月を取得"""
        now = datetime.now()
        
        # 1月前の年月を計算
        if now.month == 1:
            year = now.year - 1
            month = 12
        else:
            year = now.year
            month = now.month - 1
        
        print(f"📅 処理対象年月: {year}年{month}月（実行日の1月前）")
        return year, month
    
    def find_zip_files(self) -> List[Path]:
        """dlフォルダ内のZIPファイルを検索"""
        if not self.dl_dir.exists():
            print(f"❌ 入力ディレクトリが存在しません: {self.dl_dir}")
            return []
        
        zip_files = list(self.dl_dir.glob("*.zip"))
        print(f"📦 発見されたZIPファイル: {len(zip_files)}件")
        
        for zip_file in zip_files:
            print(f"  - {zip_file.name}")
        
        return zip_files
    
    def extract_company_name(self, subfolder_name: str, postage_pdf: Path = None, bill_pdf: Path = None) -> str:
        """会社名を抽出（シンプル版）"""
        
        # 1. bill_shippingファイルから会社名を抽出（出荷明細書の3行目を優先）
        if bill_pdf and bill_pdf.exists():
            try:
                with open(bill_pdf, 'rb') as file:
                    pdf_reader = PyPDF2.PdfReader(file)
                    
                    # すべてのページを検索して出荷明細書の3行目から会社名を抽出
                    for page_num in range(len(pdf_reader.pages)):
                        page_text = pdf_reader.pages[page_num].extract_text()
                        lines = page_text.split('\n')
                        
                        # 出荷明細書の特徴を探す（1行目に「出荷明細書」が含まれている）
                        if len(lines) > 0 and '出荷明細書' in lines[0]:
                            if len(lines) >= 3:
                                company_line = lines[2].strip()  # 3行目（0ベースなので2）
                                if '事業者名' in company_line:
                                    # 事業者名の前後をチェック
                                    parts = company_line.split('事業者名')
                                    if len(parts) > 1 and parts[1].strip():
                                        # 事業者名の後ろの文字列
                                        company_name = parts[1].strip()
                                        if company_name and len(company_name) > 1:
                                            company_name = re.sub(r'【[^】]*】', '', company_name).strip()
                                            return company_name
                                    elif len(parts) > 0 and parts[0].strip():
                                        # 事業者名の前の文字列
                                        company_name = parts[0].strip()
                                        if company_name and len(company_name) > 1:
                                            company_name = re.sub(r'【[^】]*】', '', company_name).strip()
                                            return company_name
                                elif len(lines) > 3:
                                    next_line = lines[3].strip()
                                    if next_line and len(next_line) > 2:
                                        company_name = re.sub(r'【[^】]*】', '', next_line).strip()
                                        return company_name
                            break
                    
            except:
                pass
        
        # 2. フォルダ名から会社名を抽出（最後の手段）
        if '_' in subfolder_name:
            parts = subfolder_name.split('_')
            if len(parts) > 1:
                # フォルダ名の会社コード部分をそのまま使用
                return parts[1]
        
        return '不明な会社'
    
    def split_pdf(self, pdf_path: Path, output_dir: Path, company_name: str, year: int, month: int) -> List[Path]:
        """PDFを分割（出荷明細書の場所で分割）"""
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                page_count = len(pdf_reader.pages)
                
                if page_count < 1:
                    print("⚠️ PDFが空のため分割できません")
                    return []
                
                split_files = []
                
                # 出荷明細書のページを特定
                shipping_page_index = None
                for page_num in range(page_count):
                    page_text = pdf_reader.pages[page_num].extract_text()
                    lines = page_text.split('\n')
                    if len(lines) > 0 and '出荷明細書' in lines[0]:
                        shipping_page_index = page_num
                        break
                
                if shipping_page_index is None:
                    print("⚠️ 出荷明細書ページが見つからないため、1ページ目で分割します")
                    shipping_page_index = 0
                
                # 請求書（出荷明細書ページより前）
                if shipping_page_index > 0:
                    pdf_writer = PyPDF2.PdfWriter()
                    for i in range(shipping_page_index):
                        pdf_writer.add_page(pdf_reader.pages[i])
                    
                    bill_filename = f"{year:04d}{month:02d}_{company_name}様_ご請求書.pdf"
                    bill_path = output_dir / bill_filename
                    
                    with open(bill_path, 'wb') as output_file:
                        pdf_writer.write(output_file)
                    
                    split_files.append(bill_path)
                    print(f"📁 請求書ファイル作成: {bill_filename}")
                else:
                    # 出荷明細書が1ページ目の場合、そのページを請求書として使用
                    pdf_writer = PyPDF2.PdfWriter()
                    pdf_writer.add_page(pdf_reader.pages[0])
                    
                    bill_filename = f"{year:04d}{month:02d}_{company_name}様_ご請求書.pdf"
                    bill_path = output_dir / bill_filename
                    
                    with open(bill_path, 'wb') as output_file:
                        pdf_writer.write(output_file)
                    
                    split_files.append(bill_path)
                    print(f"📁 請求書ファイル作成: {bill_filename}")
                
                # 明細書（出荷明細書ページ以降を1つのPDFにまとめる）
                pdf_writer = PyPDF2.PdfWriter()
                for i in range(shipping_page_index, page_count):
                    pdf_writer.add_page(pdf_reader.pages[i])
                
                detail_filename = f"{year:04d}{month:02d}_{company_name}様_ご請求明細書.pdf"
                detail_path = output_dir / detail_filename
                
                with open(detail_path, 'wb') as output_file:
                    pdf_writer.write(output_file)
                
                split_files.append(detail_path)
                print(f"📁 明細書ファイル作成: {detail_filename}")
                
                return split_files
                
        except Exception as e:
            print(f"❌ PDF分割エラー: {e}")
            return []
    
    def process_zip_file(self, zip_path: Path, year: int, month: int) -> bool:
        """ZIPファイルを処理"""
        try:
            print(f"\n📦 ZIPファイル処理中: {zip_path.name}")
            
            # 一時ディレクトリを作成
            self.temp_dir = Path(tempfile.mkdtemp())
            
            # ZIPファイルを解凍
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(self.temp_dir)
            
            print(f"📦 ZIP解凍完了")
            
            # 解凍されたファイルを処理
            success_count = 0
            business_count = 0
            
            for item in self.temp_dir.iterdir():
                if item.is_dir():
                    # サブフォルダ（会社フォルダ）を処理
                    business_count += 1
                    print(f"\n🏢 会社処理中 ({business_count}): {item.name}")
                    
                    # 送料明細書とbill_shippingファイルを特定
                    postage_file = None
                    bill_shipping_file = None
                    
                    for file in item.iterdir():
                        if file.is_file() and file.name.startswith('postage_') and file.name.endswith('.pdf'):
                            postage_file = file
                        elif file.is_file() and file.name.startswith('bill_shipping_') and file.name.endswith('.pdf'):
                            bill_shipping_file = file
                    
                    # 会社名を抽出
                    company_name = self.extract_company_name(item.name, postage_file, bill_shipping_file)
                    print(f"✅ 会社名抽出成功: {company_name}")
                    
                    # サブフォルダ内のファイルを処理
                    for file in item.iterdir():
                        if file.is_file():
                            if file.name.startswith('bill_shipping_') and file.name.endswith('.pdf'):
                                # bill_shippingファイルを分割
                                split_files = self.split_pdf(file, self.output_dir, company_name, year, month)
                                success_count += len(split_files)
                            
                            elif file.name.startswith('postage_') and file.name.endswith('.pdf'):
                                # 送料明細書をリネームしてコピー
                                postage_filename = f"{year:04d}{month:02d}_{company_name}様_ご請求送料明細書.pdf"
                                postage_path = self.output_dir / postage_filename
                                
                                shutil.copy2(file, postage_path)
                                print(f"📁 送料明細書ファイル作成: {postage_filename}")
                                success_count += 1
            
            print(f"✅ 処理成功: {business_count}事業者, {success_count}件のファイルを作成")
            return True
            
        except Exception as e:
            print(f"❌ ZIPファイル処理エラー: {e}")
            return False
        
        finally:
            # 一時ディレクトリを削除
            if self.temp_dir and self.temp_dir.exists():
                shutil.rmtree(self.temp_dir)
    
    def run(self):
        """メイン処理"""
        try:
            print("🚀 PDF分割・リネームツールを開始します...")
            
            # ZIPファイルを検索
            zip_files = self.find_zip_files()
            
            if not zip_files:
                print("⚠️ 処理対象のZIPファイルが見つかりません")
                print(f"💡 {self.dl_dir} にZIPファイルを配置してください")
                return
            
            # 処理対象年月を取得（実行日の1月前）
            year, month = self.get_processing_year_month()
            
            # 各ZIPファイルを処理
            success_count = 0
            total_businesses = 0
            total_files = 0
            
            for zip_file in zip_files:
                if self.process_zip_file(zip_file, year, month):
                    success_count += 1
                
                # 出力ファイル数をカウント
                output_files = list(self.output_dir.glob("*.pdf"))
                total_files = len(output_files)
            
            print(f"\n🎉 処理完了！")
            print(f"📊 処理結果:")
            print(f"   - 総ZIPファイル数: {len(zip_files)}件")
            print(f"   - 成功: {success_count}件")
            print(f"   - 失敗: {len(zip_files) - success_count}件")
            print(f"   - 出力ファイル数: {total_files}件")
            print(f"📁 出力先: {self.output_dir}")
            print(f"\n💡 次のステップ:")
            print(f"   1. {self.output_dir} 内のファイルを確認")
            print(f"   2. 必要に応じてGoogle Driveにアップロード")
            
        except Exception as e:
            print(f"❌ メイン処理エラー: {e}")

if __name__ == "__main__":
    processor = PDFProcessor()
    processor.run()