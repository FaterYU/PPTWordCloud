import WChead
# import wordcloud
import matplotlib.pyplot as plt
text = WChead.Read_ppt("D:\VScode\PPTWordCloud\demo.pptx")
# WChead.Creat_wc(text).to_file("./PPTWordCloud/demo.png")
plt.imshow(WChead.Creat_wc(text))
plt.axis("off")
plt.show()
