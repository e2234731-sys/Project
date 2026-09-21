# -*- coding: utf-8 -*-
import os, sys, subprocess, tempfile, shutil, time, uuid

if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

REPO_URL = 'https://github.com/e2234731-sys/Project.git'
PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_NAME = os.path.basename(PROJECT_DIR)

print(f'==> 开始同步项目 [{PROJECT_NAME}] 至 GitHub: {REPO_URL}...')

temp_dir = os.path.join(tempfile.gettempdir(), f'gh_sync_{uuid.uuid4().hex[:8]}')
if os.path.exists(temp_dir):
    shutil.rmtree(temp_dir, ignore_errors=True)

try:
    print('1. 正在拉取远程仓库...')
    subprocess.run(['git', 'clone', '--depth', '1', REPO_URL, temp_dir], check=True)

    target_sub = os.path.join(temp_dir, PROJECT_NAME)
    os.makedirs(target_sub, exist_ok=True)

    ignore_set = {'build', 'dist', '__pycache__', '.git', '_temp_zip_extracted_'}
    
    print('2. 正在同步最新源码与配置...')
    for root, dirs, files in os.walk(PROJECT_DIR):
        dirs[:] = [d for d in dirs if d not in ignore_set and not d.endswith('.tmp')]
        rel_path = os.path.relpath(root, PROJECT_DIR)
        dest_dir = os.path.join(target_sub, rel_path) if rel_path != '.' else target_sub
        os.makedirs(dest_dir, exist_ok=True)

        for file in files:
            if file.endswith(('.py', '.bat', '.json', '.md', '.txt', '.gitignore', '.yaml', '.yml')) and not file.startswith('~') and not file.endswith('.tmp'):
                src_file = os.path.join(root, file)
                dest_file = os.path.join(dest_dir, file)
                shutil.copy2(src_file, dest_file)

    print('3. 正在提交并推送到 GitHub...')
    subprocess.run(['git', '-C', temp_dir, 'add', '.'], check=True)
    status = subprocess.run(['git', '-C', temp_dir, 'status', '--porcelain'], capture_output=True, text=True)
    
    if status.stdout.strip():
        subprocess.run(['git', '-C', temp_dir, 'commit', '-m', 'feat(v2.5): add organize-by-group directory creation and two-way ledger synchronization'], check=True)
        subprocess.run(['git', '-C', temp_dir, 'push', 'origin', 'main'], check=True)
        print('✅ 成功推送到 GitHub 远程仓库！')
    else:
        print('✨ 远程仓库已是最新版本，无需重复提交。')

except Exception as e:
    print(f'❌ 同步失败: {e}')
finally:
    if os.path.exists(temp_dir):
        try:
            shutil.rmtree(temp_dir, ignore_errors=True)
        except Exception:
            pass
