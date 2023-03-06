from googletrans import Translator

with open('file_eng.txt') as f:
    lines = f.readlines()
    f.close()

    translator = Translator()
    for line in lines:
        line = line.replace('\n', '')
        translated = translator.translate(text = line, src="en", dest="ja")
        print(line)  # 翻訳したい文章
        print(translated.text)  # 翻訳後の文章