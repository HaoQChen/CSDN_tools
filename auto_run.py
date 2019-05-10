# -*- coding: utf-8 -*-
from git import Repo
import os
import datetime

use_github = True

def main():
    current_path = os.path.dirname(os.path.realpath(__file__))
    if use_github:
        repo = Repo(current_path)
        repo.remote().pull()#将远程内容更新到本地
    #运行相关脚本
    os.system("python3 " + current_path + "/CSDN_visited_num.py")
    if use_github:
        #得到当前日期
        now_time = int(str(datetime.datetime.now().strftime('%Y%m%d')))
        # repo.index.add("-A")
        # repo.index.commit("update " + now_time)
        repo.git.add("-A")
        repo.git.commit("-m", update " + now_time)
        repo.remote().push()

if __name__ == "__main__":
    main()