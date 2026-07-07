VERSION = "1.6.3"


WEIGHT_COLUMN_NAME = "商品重量"
INSERTED_COLUMN_NAME = "快递费"
PRICE_COLUMN_NAME = "售价"
PRODUCT_COLUMN_NAME = "商品名"
PRODUCT_CODE_COLUMN_NAME = "商品编码"
SPEC_CODE_COLUMN_NAME = "规格编码"
COMBO_PRODUCT_CODE_COLUMN_NAME = "组合商品编码"
AVAILABLE_COLUMN_NAME = "可用数"
STOCK_COLUMN_NAME = "库存"
PRODUCT_TAG_COLUMN_NAME = "商品标签"
SUPPLIER_COLUMN_NAME = "供应商"
COST_PRICE_COLUMN_NAME = "成本价"
COMBO_WEIGHT_COLUMN_NAME = "组合重量"
COMBO_COST_PRICE_COLUMN_NAME = "组合成本价"
SALES_15_DAYS_COLUMN_NAME = "15天销量"
PRICE_CACHE_FILE_NAME = "price_cache.json"
DEFAULT_PRICE_MATCH_COLUMN_NAME = PRODUCT_CODE_COLUMN_NAME
DEFAULT_PRICE_THRESHOLD = 4.5
DEFAULT_LOW_PRICE_MARGIN = 0.13
DEFAULT_HIGH_PRICE_MARGIN = 0.13
SHIPPING_FEE_RULES = [
    ("0 < 商品重量 < 0.3", 1.3),
    ("0.3 <= 商品重量 < 0.5", 1.5),
    ("0.5 <= 商品重量 < 1", 1.9),
    ("1 <= 商品重量 < 1.5", 2.3),
    ("1.5 <= 商品重量 < 2", 2.8),
    ("2 <= 商品重量 < 2.5", 3.5),
    ("2.5 <= 商品重量 < 3", 3.9),
    ("3 <= 商品重量 < 4", 4.3),
    ("4 <= 商品重量 < 5", 5.6),
    ("商品重量 >= 5", 13.5),
]
COLUMN_ALIASES = {
    COMBO_WEIGHT_COLUMN_NAME: ["组合商品重量"],
    SPEC_CODE_COLUMN_NAME: ["规则编码"],
}
DEFAULT_DELETE_RULES = [
    {
        "name": "商品名和成本价都空白",
        "conditions": [
            {"column": PRODUCT_COLUMN_NAME, "operator": "blank"},
            {"column": COST_PRICE_COLUMN_NAME, "operator": "blank"},
        ],
    },
    {
        "name": "可用数为0或空白且成本价空白",
        "conditions": [
            {"column": AVAILABLE_COLUMN_NAME, "operator": "zero_or_blank"},
            {"column": COST_PRICE_COLUMN_NAME, "operator": "blank"},
        ],
    },
    {
        "name": "商品名含下架且库存为0或空白",
        "conditions": [
            {"column": PRODUCT_COLUMN_NAME, "operator": "contains", "value": "下架"},
            {"column": STOCK_COLUMN_NAME, "operator": "zero_or_blank"},
        ],
    },
    {
        "name": "商品名含下架且可用数小于100",
        "conditions": [
            {"column": PRODUCT_COLUMN_NAME, "operator": "contains", "value": "下架"},
            {"column": AVAILABLE_COLUMN_NAME, "operator": "lt", "value": 100},
        ],
    },
    {
        "name": "商品名含zxp且可用数小于100",
        "conditions": [
            {"column": PRODUCT_COLUMN_NAME, "operator": "contains_ci", "value": "zxp"},
            {"column": AVAILABLE_COLUMN_NAME, "operator": "lt", "value": 100},
        ],
    },
    {
        "name": "商品名含微瑕品",
        "conditions": [
            {"column": PRODUCT_COLUMN_NAME, "operator": "contains", "value": "微瑕品"},
        ],
    },
    {
        "name": "可用数小于10且商品重量空白",
        "conditions": [
            {"column": AVAILABLE_COLUMN_NAME, "operator": "lt", "value": 10},
            {"column": WEIGHT_COLUMN_NAME, "operator": "blank"},
        ],
    },
    {
        "name": "可用数为0且供应商没有英文字母",
        "conditions": [
            {"column": AVAILABLE_COLUMN_NAME, "operator": "eq", "value": 0},
            {"column": SUPPLIER_COLUMN_NAME, "operator": "no_english_letter"},
        ],
    },
    {
        "name": "商品标签空白且可用数为0",
        "conditions": [
            {"column": PRODUCT_TAG_COLUMN_NAME, "operator": "blank"},
            {"column": AVAILABLE_COLUMN_NAME, "operator": "eq", "value": 0},
        ],
    },
    {
        "name": "商品标签含家具或赠品",
        "any": [
            {"column": PRODUCT_TAG_COLUMN_NAME, "operator": "contains", "value": "家具"},
            {"column": PRODUCT_TAG_COLUMN_NAME, "operator": "contains", "value": "赠品"},
        ],
    },
    {
        "name": "商品标签含待清理且可用数小于30",
        "conditions": [
            {"column": PRODUCT_TAG_COLUMN_NAME, "operator": "contains", "value": "待清理"},
            {"column": AVAILABLE_COLUMN_NAME, "operator": "lt", "value": 30},
        ],
    },
    {
        "name": "商品标签含分销商且可用数小于100",
        "conditions": [
            {"column": PRODUCT_TAG_COLUMN_NAME, "operator": "contains", "value": "分销商"},
            {"column": AVAILABLE_COLUMN_NAME, "operator": "lt", "value": 100},
        ],
    },
    {
        "name": "商品标签4位数字小于2601且可用数小于20且15天销量小于15",
        "conditions": [
            {"column": PRODUCT_TAG_COLUMN_NAME, "operator": "four_digit_lt", "value": 2601},
            {"column": AVAILABLE_COLUMN_NAME, "operator": "lt", "value": 20},
            {"column": SALES_15_DAYS_COLUMN_NAME, "operator": "lt", "value": 15},
        ],
    },
    {
        "name": "供应商和成本价都空白",
        "conditions": [
            {"column": SUPPLIER_COLUMN_NAME, "operator": "blank"},
            {"column": COST_PRICE_COLUMN_NAME, "operator": "blank"},
        ],
    },
]
