#!/usr/bin/env python3
"""
PDF分割・リネームツール
dlフォルダ内のZIPファイルを処理して、outputフォルダに結果を出力
"""

import os
import sys
import zipfile
import tempfile
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
    
    def get_year_month_from_zip(self, zip_files: List[Path]) -> Tuple[int, int]:
        """ZIPファイル名から処理対象年月を自動取得"""
        try:
            for zip_file in zip_files:
                # ZIPファイル名から年月を抽出
                # 例: 請求書・出荷明細_2025-08-01～2025-08-31_202509031219.zip
                match = re.search(r'(\d{4})-(\d{2})', zip_file.name)
                if match:
                    year = int(match.group(1))
                    month = int(match.group(2))
                    print(f"📅 ZIPファイル名から年月を取得: {year}年{month}月")
                    return year, month
            
            # 抽出できない場合は現在の日付を使用
            now = datetime.now()
            print(f"⚠️ ZIPファイル名から年月を抽出できません。現在の日付を使用: {now.year}年{now.month}月")
            return now.year, now.month
            
        except Exception as e:
            print(f"❌ 年月取得でエラー: {e}")
            now = datetime.now()
            return now.year, now.month
    
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
        """PDF内容から会社名を抽出（送料明細書の事業者名を優先）"""
        # 1. 送料明細書から事業者名を抽出
        if postage_pdf and postage_pdf.exists():
            try:
                company_name = self.extract_company_name_from_postage(postage_pdf)
                if company_name and company_name != '不明な会社':
                    return company_name
            except Exception as e:
                print(f"⚠️ 送料明細書からの会社名抽出でエラー: {e}")
        
        # 2. bill_shippingファイルから会社名を抽出
        if bill_pdf and bill_pdf.exists():
            try:
                company_name = self.extract_company_name_from_pdf(bill_pdf)
                if company_name and company_name != '不明な会社':
                    return company_name
            except Exception as e:
                print(f"⚠️ bill_shippingからの会社名抽出でエラー: {e}")
        
        return '不明な会社'
    
    def extract_company_name_from_postage(self, pdf_path: Path) -> str:
        """送料明細書から事業者名を抽出"""
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                
                # 全ページからテキストを抽出
                full_text = ""
                for page in pdf_reader.pages:
                    full_text += page.extract_text() + "\n"
                
                # 事業者名の横にある文字列をそのまま取得
                lines = full_text.split('\n')
                for line in lines:
                    if '事業者名' in line:
                        # 事業者名の行から会社名を抽出
                        parts = line.split('事業者名')
                        if len(parts) > 1:
                            company_name = parts[1].strip()
                            if company_name and len(company_name) > 2:
                                # 【〜】の部分を削除
                                company_name = re.sub(r'【[^】]*】', '', company_name).strip()
                                print(f"📄 送料明細書から事業者名を抽出: {company_name}")
                                return company_name
                
                # 事業者名が見つからない場合は、一般的な会社名パターンを検索
                return self.extract_company_name_from_pdf(pdf_path)
                
        except Exception as e:
            print(f"❌ 送料明細書内容抽出エラー: {e}")
            return '不明な会社'
    
    def extract_company_name_from_pdf(self, pdf_path: Path) -> str:
        """PDF内容から会社名を抽出"""
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                
                # 最初の数ページからテキストを抽出
                for page_num in range(min(3, len(pdf_reader.pages))):
                    page = pdf_reader.pages[page_num]
                    text = page.extract_text()
                    
                    # 会社名のパターンを検索（より厳密なパターン）
                    company_patterns = [
                        r'([一-龯ァ-ヶー]{2,}株式会社)',
                        r'([一-龯ァ-ヶー]{2,}有限会社)',
                        r'([一-龯ァ-ヶー]{2,}合資会社)',
                        r'([一-龯ァ-ヶー]{2,}合名会社)',
                        r'([一-龯ァ-ヶー]{2,}協同組合)',
                        r'([一-龯ァ-ヶー]{2,}協会)',
                        r'([一-龯ァ-ヶー]{2,}財団)',
                        r'([一-龯ァ-ヶー]{2,}法人)',
                    ]
                    
                    for pattern in company_patterns:
                        matches = re.findall(pattern, text)
                        if matches:
                            # 最も長い会社名を選択
                            company_name = max(matches, key=len).strip()
                            
                            if len(company_name) > 3 and len(company_name) < 50:  # 適切な長さチェック
                                print(f"📄 PDF内容から会社名を抽出: {company_name}")
                                return company_name
                
                return '不明な会社'
                
        except Exception as e:
            print(f"❌ PDF内容抽出エラー: {e}")
            return '不明な会社'
    
    def split_pdf(self, pdf_path: Path, output_dir: Path, company_name: str, year: int, month: int) -> List[Path]:
        """PDFを分割"""
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                page_count = len(pdf_reader.pages)
                
                print(f"📄 PDFページ数: {page_count}")
                
                if page_count < 2:
                    print("⚠️ PDFが1ページのみのため分割できません")
                    return []
                
                split_files = []
                
                # 1ページ目（請求書）
                pdf_writer = PyPDF2.PdfWriter()
                pdf_writer.add_page(pdf_reader.pages[0])
                
                bill_filename = f"{year:04d}{month:02d}_{company_name}様_ご請求書.pdf"
                bill_path = output_dir / bill_filename
                
                with open(bill_path, 'wb') as output_file:
                    pdf_writer.write(output_file)
                
                split_files.append(bill_path)
                print(f"📁 請求書ファイル作成: {bill_filename}")
                
                # 2ページ目以降（明細書）
                if page_count > 1:
                    pdf_writer = PyPDF2.PdfWriter()
                    for i in range(1, page_count):
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
            print(f"📁 一時ディレクトリ: {self.temp_dir}")
            
            # ZIPファイルを解凍
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(self.temp_dir)
            
            print(f"📦 ZIP解凍完了")
            
            # 解凍されたファイルを処理
            success_count = 0
            
            for item in self.temp_dir.iterdir():
                if item.is_dir():
                    # サブフォルダ（会社フォルダ）を処理
                    print(f"\n🏢 会社処理中: {item.name}")
                    
                    # まず送料明細書から会社名を抽出（事業者名項目）
                    postage_file = None
                    bill_shipping_file = None
                    
                    for file in item.iterdir():
                        if file.is_file() and file.name.startswith('postage_') and file.name.endswith('.pdf'):
                            postage_file = file
                        elif file.is_file() and file.name.startswith('bill_shipping_') and file.name.endswith('.pdf'):
                            bill_shipping_file = file
                    
                    # 会社名を抽出（送料明細書を優先、次にbill_shipping）
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
                                
                                import shutil
                                shutil.copy2(file, postage_path)
                                print(f"📁 送料明細書ファイル作成: {postage_filename}")
                                success_count += 1
            
            print(f"✅ 処理成功: {success_count}件のファイルを作成")
            return True
            
        except Exception as e:
            print(f"❌ ZIPファイル処理エラー: {e}")
            return False
        
        finally:
            # 一時ディレクトリを削除
            if self.temp_dir and self.temp_dir.exists():
                import shutil
                shutil.rmtree(self.temp_dir)
                print("🧹 一時ファイルを削除しました")
    
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
            
            # ZIPファイル名から年月を取得
            year, month = self.get_year_month_from_zip(zip_files)
            print(f"📅 処理対象年月: {year}年{month}月")
            
            # 各ZIPファイルを処理
            success_count = 0
            for zip_file in zip_files:
                if self.process_zip_file(zip_file, year, month):
                    success_count += 1
            
            print(f"\n🎉 処理完了！")
            print(f"📊 処理結果: 総ZIPファイル数: {len(zip_files)}件, 成功: {success_count}件, 失敗: {len(zip_files) - success_count}件")
            print(f"📁 出力先: {self.output_dir}")
            print(f"\n💡 次のステップ:")
            print(f"   1. {self.output_dir} 内のファイルを確認")
            print(f"   2. 必要に応じてGoogle Driveにアップロード")
            
        except Exception as e:
            print(f"❌ メイン処理エラー: {e}")

if __name__ == "__main__":
    processor = PDFProcessor()
    processor.run()
