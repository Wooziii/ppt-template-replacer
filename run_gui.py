from __future__ import annotations

import argparse

from prototype_tool.gui import main


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="PPT 模板替换桌面界面")
    parser.add_argument("--smoke-test", action="store_true", help="只验证窗口能否成功创建")
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    main(smoke_test=args.smoke_test)
