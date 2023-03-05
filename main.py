import win32com.client
import time
import sys

spk = win32com.client.Dispatch("SAPI.SpVoice")

f = open("article.txt", "r", encoding="utf-8")
all_content = f.read().replace("\n", " ")
f.close()

len_all_words = len(all_content.split(" "))

Separators = [".", ",", "!", "?"]
Separators_Corresponding_Letter = {".": "句号", ",": "逗号", "!": "感叹号", "?": "问号"}

Processed_contents = []
Processed_Part = ""

for letter in all_content:
    Processed_Part += letter
    if letter in Separators:
        Processed_contents.append(Processed_Part)
        Processed_Part = ""

del Processed_Part

f = open("Processed.txt", "w")
for l in Processed_contents:
    f.write(l)
f.close()

Time_Should_Use = str(len(all_content) * 0.5 + len(Processed_contents) * 1) + "s"

print("总词量", len_all_words, "| 句子数量", len(Processed_contents), "| 预计完成时间", Time_Should_Use)
print("已完成格式化，倒数 3 秒准备开始！")
time.sleep(3)

def printp(c):
    sys.stdout.write(c)
    sys.stdout.flush()

def read_word(word, sleep_time:int=0.5):
    printp(" ")

    for w in word:
        printp(w)
        time.sleep(0.05)
    
    spk.Speak(word)
    time.sleep(len(word) * (sleep_time-0.05))

    if word[-1] in Separators:
        spk.Speak(Separators_Corresponding_Letter[word[-1]])
        time.sleep(1)

for content in Processed_contents:
    if not content:
        continue

    words = content.split(" ")

    for word in words:
        if word:
            read_word(word)

print("\n[ OK ] 文章已完成！")
spk.Speak("文章已完成！")
sys.exit()