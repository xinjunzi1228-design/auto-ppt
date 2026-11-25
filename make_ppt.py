from pptx import Presentation
from pptx.util import Inches, Pt
from datetime import datetime

prs = Presentation()
prs.slide_width = Inches(13.333)   # 16:9
prs.slide_height = Inches(7.5)

# 封面
s = prs.slides.add_slide(prs.slide_layouts[0])
s.shapes.title.text = "自动生成的演示文稿"
s.placeholders[1].text = f"生成日期：{datetime.now():%Y-%m-%d}"

# 内容页
s = prs.slides.add_slide(prs.slide_layouts[1])
s.shapes.title.text = "要点"
tf = s.placeholders[1].text_frame
tf.text = "第一点：由 GitHub Actions 自动生成"
for line in ["第二点：可定时运行", "第三点：可手动触发"]:
    p = tf.add_paragraph()
    p.text = line
    p.font.size = Pt(24)

prs.save("output.pptx")
print("PPT 已生成")
