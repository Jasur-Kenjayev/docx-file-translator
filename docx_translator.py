from docx import Document
from googletrans import Translator
from tqdm import tqdm
import random
import os

os.system("cls")
red    = "\033[31m"
blue   = "\033[34m"
bold   = "\033[1m"
reset  = "\033[0m"
green  = "\033[32m"
yellow = "\033[33m"
colors = [
    "\033[38;5;226m",
    "\033[38;5;227m",
    "\033[38;5;229m",
    "\033[38;5;230m",
    "\033[38;5;190m",
    "\033[38;5;191m",
    "\033[38;5;220m",
    "\033[38;5;221m",
    "\033[38;5;142m",
    "\033[38;5;214m",
]

color1, color2, color3, color4, color5 = random.sample(colors, 5)
baner = f"""\033[33m
   ____                     ______                      __      __
   / __ \____  ______  __   /_  __/________ _____  _____/ /___ _/ /_____  _____
  / / / / __ \/ ___/ |/_/    / / / ___/ __ `/ __ \/ ___/ / __ `/ __/ __ \/ ___/
 / /_/ / /_/ / /___>  <     / / / /  / /_/ / / / (__  ) / /_/ / /_/ /_/ / /
/_____/\____/\___/_/|_|    /_/ /_/   \__,_/_/ /_/____/_/\__,_/\__/\____/_/
                                                       

                      ᴰᵉᵛᵉˡᵒᵖᵉʳ ᵀᵉˡᵉᵍʳᵃᵐ: @ᴾʸᵗʰᵒⁿ_ᴷᵒᵈᵉʳˢ
                            ᵖʰᵒⁿᵉ: +⁹⁹⁸³³³⁰⁰⁹⁸⁸⁸
                      \033[32mhttps://github.com/Jasur-Kenjayev
            {color1}[\033[31m#\033[0m\033[33m]\033[31mProgram information: This program translates word file
"""
print(baner)

docx_input = input("\033[34mEnter Word file: ")

input_language = input("\n\033[32mEnter the language code: ")

def rasmli_docx_tarjima_qilish(docx_fayl_nomi, yangi_fayl_nomi, til_kodi):
    doc = Document(docx_fayl_nomi)
    translator = Translator()

    for paragraph in tqdm(doc.paragraphs):
        ozbek_matn = paragraph.text
        for run in paragraph.runs:
            rasm_usti_matn = run.text
            try:
                ingliz_matn = translator.translate(rasm_usti_matn, dest=til_kodi).text
                run.text = ingliz_matn

            except Exception as e:
                pass
    print("\n\nTranslation Completed Successfully...")
    print(f"\n\033[31mFile Saved as {yangi_fayl_nomi}")

    doc.save(yangi_fayl_nomi)

id = random.randint(1,10000)
docx_fayl_nomi = docx_input
yangi_fayl_nomi = f"{input_language}_tarjima{id}.docx"
til_kodi = input_language

rasmli_docx_tarjima_qilish(docx_fayl_nomi, yangi_fayl_nomi, til_kodi)