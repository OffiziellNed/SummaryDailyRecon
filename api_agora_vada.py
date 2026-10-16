from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from curl_cffi import requests
from bs4 import BeautifulSoup

app = FastAPI(title="Agora Vada API - Anti-Bot Edition")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class URLRequest(BaseModel):
    url: str

@app.post("/api/tarik-berita")
def tarik_berita(req: URLRequest):
    try:
        # Senjata Utama: impersonate="chrome110" meniru sidik jari Google Chrome
        response = requests.get(
            req.url, 
            impersonate="chrome110",
            timeout=15
        )
        
        if response.status_code != 200:
            raise HTTPException(status_code=response.status_code, detail="Akses tetap ditolak oleh web.")

        soup = BeautifulSoup(response.content, 'html.parser')
        
        # Ekstraksi Judul
        title_tag = soup.find('title')
        title = title_tag.text.strip() if title_tag else "Tanpa Judul"
        
        # Ekstraksi Full Text (seperti logika JavaScript sebelumnya)
        article_content = ""
        for p in soup.find_all('p'):
            text = p.get_text(strip=True)
            if len(text) > 40:
                article_content += text + "\n\n"

        if not article_content.strip():
            meta_desc = soup.find('meta', attrs={'name': 'description'})
            article_content = meta_desc['content'] if meta_desc else "Deskripsi tidak ditemukan."

        return {
            "status": "success",
            "title": title,
            "description": article_content.strip(),
            "prompt": f"Judul: {title}\n\nIsi Berita Lengkap:\n{article_content.strip()}"
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Gagal menyedot: {str(e)}")
