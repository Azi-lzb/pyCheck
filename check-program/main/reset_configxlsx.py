from build_strategy1 import CONFIG_FILE, initialize_config_workbook


def main():
    rows = initialize_config_workbook(CONFIG_FILE)
    print(f"config.xlsx 已初始化完成，共写入 {len(rows)} 行策略一基准配置。")


if __name__ == "__main__":
    main()
