"""
同事专用入口脚本（只读）

用途：
1) 验证只读连接
2) 查看收件箱总量
3) 搜索最近 N 天、邮件头含指定关键词（默认 911600210）
4) 打印邮件摘要 + 正文预览（不改邮件状态）
"""

from readonly_mail_client import ReadonlyFareportMailClient
from secure_config import get_config

# 在邮件头（主题/发件人/日期）中匹配，例如产品编号出现在主题里
FILTER_KEYWORDS = ["911600210"]
SEARCH_DAYS = 2


def main() -> None:
    print("=" * 60)
    print("同事协作模式：只读邮箱分析")
    print("=" * 60)

    config = get_config()
    client = ReadonlyFareportMailClient(
        email_address=config["email_address"],
        password=config["password"],
        imap_server=config.get("imap_server", "imap.exmail.qq.com"),
        imap_port=int(config.get("imap_port", 993)),
    )

    if not client.connect():
        print("❌ 连接失败，请检查环境变量或网络")
        return

    print("✅ 只读连接成功")
    print(f"📧 收件箱总邮件数：{client.get_mail_count()}")

    mail_ids = client.search_emails(
        days=SEARCH_DAYS,
        unread=False,
        subject_keywords=FILTER_KEYWORDS,
    )
    print(f"🔎 最近 {SEARCH_DAYS} 天、含关键词 {FILTER_KEYWORDS}：{len(mail_ids)} 封")

    show_count = min(20, len(mail_ids))
    if show_count > 0:
        print(f"\n显示前 {show_count} 封摘要：\n")
    for i, mail_id in enumerate(mail_ids[:show_count], 1):
        summary = client.fetch_email_summary(mail_id)
        if not summary:
            continue
        print(f"[{i:2}] ID: {summary['id']}")
        print(f"     日期：{summary['date']}")
        print(f"     发件人：{summary['from']}")
        print(f"     主题：{summary['subject']}")
        preview = client.fetch_email_body_preview(mail_id, max_chars=180)
        if preview:
            print(f"     正文预览：{preview}")
        print()

    client.disconnect()
    print("✅ 已断开连接")


if __name__ == "__main__":
    main()
