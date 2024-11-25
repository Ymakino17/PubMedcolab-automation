# run_colab_notebook.py
import os
from google.colab import drive

# Google Driveをマウント
drive.mount('/content/drive')

# ノートブックを実行するコマンド
!jupyter nbconvert --to notebook --execute "PubMed_monitoring.ipynb"
