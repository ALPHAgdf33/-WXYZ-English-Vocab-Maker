# -WXYZ-English-Vocab-Maker
A PPT maker, using Python and Google TTS.

Copy the following content to the terminal to download the frontend library.
```
pip install python-pptx python-docx requests beautifulsoup4 gTTS
```

You can just copy the source code below for the most simple usage.
Try with the EnglishVocab docx file.
```py
import re
import io
import json
import requests
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from gtts import gTTS
from bs4 import BeautifulSoup


def get_image_stream(query):
    """从 Bing 爬取参考图片"""
    search_url = f"https://www.bing.com/images/search?q={query}+meaning+illustration&form=HDRSC2"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    try:
        response = requests.get(search_url, headers=headers, timeout=5)
        soup = BeautifulSoup(response.text, 'html.parser')
        m_tags = soup.find_all("a", {"class": "iusc"})
        if m_tags:
            img_url = json.loads(m_tags[0].get("m")).get("murl")
            img_res = requests.get(img_url, timeout=5)
            return io.BytesIO(img_res.content)
    except:
        return None


def set_mixed_font(paragraph, text, size, english_font='Times New Roman', is_bold=False, color_rgb=(0, 0, 0)):
    """处理中英文混排：中文用微软雅黑，英文用指定字体"""
    for char in text:
        run = paragraph.add_run()
        run.text = char
        run.font.size = size
        run.font.bold = is_bold
        run.font.color.rgb = RGBColor(*color_rgb)
        if '\u4e00' <= char <= '\u9fff' or char in '，。？！：；（）“”‘’':
            run.font.name = 'Microsoft YaHei UI'
        else:
            run.font.name = english_font


def create_final_ai_ppt(docx_path, output_ppt="Advanced_Vocab_AI.pptx"):
    # 读取 Word 文档 [cite: 1]
    doc = Document(docx_path)
    full_text = "\n".join([para.text for para in doc.paragraphs if para.text.strip()])

    # 正则表达式：匹配单词、音标（含多音标）、释义、同义词、例句
    pattern = r"(\d+)\.(.*?)\s(/.*?/.*?)\s(.*?)\n\[同义词\](.*?)\n(.*?)\n(.*?)(?=\n\d+\.|\Z)"
    matches = re.findall(pattern, full_text, re.S)

    prs = Presentation()

    for match in matches:
        index, word, phonetics, definition, synonyms, sentence_en, sentence_zh = match
        word_clean = word.strip()

        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # --- 1. 标题 (Consolas) ---
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(5), Inches(1))
        set_mixed_font(title_box.text_frame.paragraphs[0], word_clean, Pt(42), english_font='Consolas', is_bold=True,
                       color_rgb=(0, 51, 102))

        # --- 2. 音标与释义 (Times New Roman + 微软雅黑) ---
        info_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(5), Inches(2.5))
        tf = info_box.text_frame
        tf.word_wrap = True
        set_mixed_font(tf.paragraphs[0], phonetics.strip(), Pt(18), color_rgb=(128, 128, 128))
        set_mixed_font(tf.add_paragraph(), f"释义: {definition.strip()}", Pt(24), is_bold=True)
        set_mixed_font(tf.add_paragraph(), f"同义词: {synonyms.strip()}", Pt(16))

        # --- 3. 例句 ---
        ex_box = slide.shapes.add_textbox(Inches(0.5), Inches(4.2), Inches(5), Inches(3))
        tf_ex = ex_box.text_frame
        tf_ex.word_wrap = True
        set_mixed_font(tf_ex.paragraphs[0], sentence_en.strip(), Pt(20))
        set_mixed_font(tf_ex.add_paragraph(), sentence_zh.strip(), Pt(16), color_rgb=(80, 80, 80))

        # --- 4. 插入图片 ---
        print(f"正在处理: {word_clean}")
        img_stream = get_image_stream(word_clean)
        if img_stream:
            try:
                slide.shapes.add_picture(img_stream, Inches(5.8), Inches(1.2), height=Inches(3.5))
            except:
                pass

        # --- 5. AI 语音生成与插入 ---
        try:
            # 朗读内容：单词 + 停顿 + 例句
            audio_text = f"{word_clean}. ... {sentence_en.strip()}"
            tts = gTTS(text=audio_text, lang='en')
            audio_stream = io.BytesIO()
            tts.write_to_fp(audio_stream)
            audio_stream.seek(0)

            # 插入音频图标到右上角
            slide.shapes.add_movie(
                audio_stream, Inches(8.5), Inches(0.3), Inches(0.5), Inches(0.5),
                mime_type='audio/mp3'
            )
        except Exception as e:
            print(f"语音生成失败 ({word_clean}): {e}")

    prs.save(output_ppt)
    print(f"\n成功！已生成包含语音和图片的 PPT: {output_ppt}")


# 执行脚本
create_final_ai_ppt('高级词汇（共61组）.docx')
```
