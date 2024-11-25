import json
import requests
import xml.etree.ElementTree as ET
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta
from openpyxl import Workbook, load_workbook
import os

# ファイル名設定
keyword_file = "search_keywords.json"
result_file = "PubMed_results.xlsx"

# PubMed API設定
api_key = "a7958d4ab8f82c1de8158c70b276d935b908"  # PubMedのAPIキー
recipient = "yasu1986m@gmail.com"  # メール受信者
sender_email = "yasu1986m@gmail.com"  # 送信者メールアドレス
sender_password = "edsu fxnd gqss liad"  # 送信者メールパスワード

# キーワードの保存
def save_keywords(keywords):
    with open(keyword_file, "w") as f:
        json.dump(keywords, f)

# キーワードの読み込み
def load_keywords():
    try:
        with open(keyword_file, "r") as f:
            return json.load(f)
    except FileNotFoundError:
        return ["RNA splicing"]

# PubMed API検索
def search_pubmed(query, start_date, end_date, api_key):
    base_url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
    date_range = f"{start_date}:{end_date}[dp]"
    full_query = f"{query} AND {date_range}"
    params = {
        "db": "pubmed",
        "term": full_query,
        "retmode": "json",
        "sort": "date",
        "retmax": 10,
        "api_key": api_key
    }
    response = requests.get(base_url, params=params)
    response.raise_for_status()
    return response.json()["esearchresult"]["idlist"]

# PubMed APIで論文詳細を取得
def fetch_abstracts(ids, api_key):
    base_url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi"
    params = {
        "db": "pubmed",
        "id": ",".join(ids),
        "retmode": "xml",
        "rettype": "abstract",
        "api_key": api_key
    }
    response = requests.get(base_url, params=params)
    response.raise_for_status()
    return response.text

# XML解析
def parse_pubmed_data(xml_data):
    root = ET.fromstring(xml_data)
    articles = []
    for article in root.findall(".//PubmedArticle"):
        pmid = article.find(".//PMID").text if article.find(".//PMID") is not None else "No PMID"
        title = article.find(".//ArticleTitle").text if article.find(".//ArticleTitle") is not None else "No Title"
        authors = ", ".join(
            f"{author.find('ForeName').text} {author.find('LastName').text}"
            for author in article.findall(".//Author")
            if author.find("ForeName") is not None and author.find("LastName") is not None
        )
        journal = article.find(".//Title").text if article.find(".//Title") is not None else "No Journal"
        pub_date = article.find(".//PubDate/Year").text if article.find(".//PubDate/Year") is not None else "No Date"
        abstract_parts = [abstract.text for abstract in article.findall(".//AbstractText") if abstract.text]
        abstract = " ".join(abstract_parts) if abstract_parts else "No Abstract"
        link = f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/"

        articles.append({
            "pmid": pmid,
            "title": title,
            "authors": authors,
            "journal": journal,
            "pub_date": pub_date,
            "abstract": abstract,
            "link": link,
        })
    return articles

# Excel保存
def save_to_excel(articles, sheet_name):
    if os.path.exists(result_file):
        workbook = load_workbook(result_file)
    else:
        workbook = Workbook()
    if sheet_name in workbook.sheetnames:
        del workbook[sheet_name]
    sheet = workbook.create_sheet(sheet_name)
    sheet.append(["PMID", "Title", "Authors", "Journal", "Publication Date", "Abstract", "Link"])
    for article in articles:
        sheet.append([article["pmid"], article["title"], article["authors"], article["journal"], article["pub_date"], article["abstract"], article["link"]])
    workbook.save(result_file)

# メール送信
def send_email(subject, html_content):
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = sender_email
    msg["To"] = recipient
    msg.attach(MIMEText(html_content, "html"))
    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipient, msg.as_string())

# メイン処理
def main():
    keywords = load_keywords()
    start_date = (datetime.now() - timedelta(days=1)).strftime("%Y/%m/%d")
    end_date = datetime.now().strftime("%Y/%m/%d")
    for keyword in keywords:
        ids = search_pubmed(keyword, start_date, end_date, api_key)
        if ids:
            xml_data = fetch_abstracts(ids, api_key)
            articles = parse_pubmed_data(xml_data)
            save_to_excel(articles, sheet_name=keyword)
            html_content = f"<html><body><h1>Results for {keyword}</h1></body></html>"
            send_email(subject=f"PubMed Results - {keyword}", html_content=html_content)
        else:
            print(f"No new articles for keyword: {keyword}")

if __name__ == "__main__":
    main()
