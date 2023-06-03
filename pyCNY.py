import pinyin
from openpyxl import Workbook

# Create a new workbook
workbook = Workbook()

# Create a new sheet for each character
sheet = workbook.create_sheet(title='cnyChar')

# Sample input
Alst = "爱 八 爸 杯 本 不 菜 茶 车 吃 打 大 的 点 电 店 都 读 对 多 儿 二 饭 飞 分 高 个 工 国 果 汉 好 号 喝 和 很 后 话 回 会 机 几 家 见 叫 她 姐 今 九 觉 开 看 客 块 来 老 冷 里 了 六 妈 吗 买 么 没 们 米 面 名 明 哪 那 呢 能 你 年 女 朋 七 气 前 钱 请 去 热 人 认 三 商 上 少 生 师 十 什 时 识 视 是 书 谁 水 睡 说 四 岁 他 太 天 听 同 我 五 午 下 先 现 想 小 校 些 写 谢 学 样 一 衣 医 椅 影 友 有 雨 语 院 月 再 在 怎 这 中 钟 住 桌 子 字 昨 作 坐 做"
Blst = Alst.replace(" ", "")
lst = list(Blst)

counter = 0
# Loop through rows
for row in range(1, 100, 3):
    # Loop through columns (A to T)
    for col in range(1, 21):
        if counter == len(lst):
            break
        cell1 = sheet.cell(row=row, column=col)
        cell2 = sheet.cell(row=row + 1, column=col)
        # print(cell1.coordinate, cell2.coordinate)

        sheet[str(cell1.coordinate)] = lst[counter]
        sheet[str(cell2.coordinate)] = pinyin.get(lst[counter])
        counter += 1

# Remove the default sheet created initially
workbook.remove(workbook["Sheet"])

# Save the workbook
workbook.save("output.xlsx")