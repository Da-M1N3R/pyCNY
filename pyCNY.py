import pinyin
from openpyxl import Workbook
import func

# Create a new workbook
workbook = Workbook()

# Create a new sheet for each character
sheet = workbook.create_sheet(title='cnyChar')

# Sample input
hsk1 = "爱 八 爸 杯 本 不 菜 茶 车 吃 打 大 的 点 电 店 都 读 对 多 儿 二 饭 飞 分 高 个 工 国 果 汉 好 号 喝 和 很 后 话 回 会 机 几 家 见 叫 她 姐 今 九 觉 开 看 客 块 来 老 冷 里 了 六 妈 吗 买 么 没 们 米 面 名 明 哪 那 呢 能 你 年 女 朋 七 气 前 钱 请 去 热 人 认 三 商 上 少 生 师 十 什 时 识 视 是 书 谁 水 睡 说 四 岁 他 太 天 听 同 我 五 午 下 先 现 想 小 校 些 写 谢 学 样 一 衣 医 椅 影 友 有 雨 语 院 月 再 在 怎 这 中 钟 住 桌 子 字 昨 作 坐 做"


hsk2 = "吧 白 百 班 帮 报 比 笔 边 便 表 别 病 步 长 常 场 唱 出 穿 床 次 从 错 但 到 道 得 等 弟 第 懂 动 房 非 告 哥 歌 给 公 共 狗 瓜 贵 过 孩 黑 红 欢 还 火 鸡 间 件 教 近 进 经 就 考 可 课 快 乐 累 离 两 路 卖 慢 忙 猫 每 妹 门 奶 男 您 牛 旁 跑 票 起 汽 千 情 晴 球 然 让 日 肉 色 身 事 试 室 手 司 思 送 诉 虽 所 它 题 体 跳 外 完 玩 晚 往 为 问 洗 笑 新 姓 休 雪 眼 羊 药 要 也 已 以 意 因 阴 泳 游 右 鱼 员 远 运 早 站 找 真 正 知 纸 助 着 走 最 左"


lst = func.setUpLst(hsk2)

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
workbook.save("output2.xlsx")