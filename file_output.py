"""Publish completed output without replacing an existing file."""
import os
import shutil
import tempfile
from pathlib import Path


def publish_file(completed: Path, directory: Path, stem: str, suffix: str) -> Path:
    directory = directory.resolve(strict=True)
    if not directory.is_dir():
        raise ValueError('输出位置不是文件夹。')
    safe_stem = stem[:100].rstrip(' .') or '文件'
    fd, staging_name = tempfile.mkstemp(prefix='.ruyi-', suffix='.tmp', dir=directory)
    staging = Path(staging_name)
    try:
        with os.fdopen(fd, 'wb') as output, completed.open('rb') as source:
            shutil.copyfileobj(source, output)
            output.flush()
            os.fsync(output.fileno())
        for index in range(1, 10001):
            extra = '' if index == 1 else f'_{index}'
            target = directory / f'{safe_stem}_如意{extra}{suffix}'
            try:
                if os.name == 'nt':
                    # On Windows rename fails atomically when the destination exists.
                    os.rename(staging, target)
                else:
                    os.link(staging, target)
                from desktop_integration import notify_shell
                notify_shell(target)
                return target
            except FileExistsError:
                continue
        raise ValueError('同名输出过多，请选择其他文件夹。')
    finally:
        staging.unlink(missing_ok=True)
