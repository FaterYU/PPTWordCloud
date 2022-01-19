import imp
import WChead
# import wordcloud
import matplotlib.pyplot as plt
import sys
import os
# text = WChead.Read_ppt("D:\VScode\PPTWordCloud\demo.pptx")
# # WChead.Creat_wc(text).to_file("./PPTWordCloud/demo.png")
# plt.imshow(WChead.Creat_wc(text))
# plt.axis("off")
# plt.show()

if __name__ == '__main__':

    path = sys.path[0]
    print(path)

    for filename in os.listdir(path):
        if bool(filename.find(".pptx")+1):
            # print(filename)
            # print(path + "\\" + filename)
            # print(path+"\\"+filename+".png")
            text = WChead.Read_ppt(path + "\\" + filename)
            WChead.Creat_wc(text).to_file(path+"\\"+filename+".png")
    # text = WChead.Read_ppt()
