import os
# 强行在代码级别设置镜像站，无视终端环境
os.environ["HF_ENDPOINT"] = "https://hf-mirror.com"
from huggingface_hub import snapshot_download

print("🚀 开始通过国内镜像下载模型...")
snapshot_download(
    repo_id="sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2",
    local_dir="D:/hf_models/manual_download/",
    local_dir_use_symlinks=False
)
print("✅ 下载完成！请去 D 盘查看。")