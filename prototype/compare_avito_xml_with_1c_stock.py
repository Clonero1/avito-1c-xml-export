from __future__ import annotations

import argparse
import json
import re
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pandas as pd
from lxml import etree


CYRILLIC_TO_LATIN = str.maketrans(
    {
        "А": "A",
        "В": "B",
        "Е": "E",
        "К": "K",
        "М": "M",
        "Н": "H",
        "О": "O",
        "Р": "P",
        "С": "C",
        "Т": "T",
        "Х": "X",
        "У": "Y",
    }
)

DEFAULT_COMPARISON = {
    "strip_rn_suffix": True,
    "remove_leading_zeroes_for_numeric_codes": True,
    "use_catalog_number_2": True,
    "aggregate_stock_by_1c_code": True,
}

DEFAULT_XML_OPTIONS = {
    "update_price": False,
    "preserve_order": True,
}


@dataclass
class StockItem:
    code_1c: str = ""
    article: str = ""
    catalog_number_2: str = ""
    producer: str = ""
    alpha_code: str = ""
    nomenclature_guid: str = ""
    name: str = ""
    qty: float = 0
    price: Any = None
    keys: list[str] = field(default_factory=list)
    row_numbers: list[int] = field(default_factory=list)


@dataclass
class XmlAd:
    index: int
    ad_id: str = ""
    avito_id: str = ""
    oem: str = ""
    title: str = ""
    price: str = ""
    manager_name: str = ""
    contact_phone: str = ""
    keys: list[str] = field(default_factory=list)
    node: Any = None


def normalize_part(
    value: Any,
    *,
    remove_leading_zeroes_for_numeric_codes: bool = True,
) -> str:
    """Return one normalized comparison key."""
    if value is None or pd.isna(value):
        return ""

    text = str(value).strip().upper()
    if not text:
        return ""

    text = text.translate(CYRILLIC_TO_LATIN)
    text = re.sub(r"\([^)]*\)", "", text)
    text = re.sub(r"[^A-Z0-9]", "", text)

    if (
        remove_leading_zeroes_for_numeric_codes
        and text.isdigit()
        and len(text) > 1
    ):
        stripped = text.lstrip("0")
        if stripped:
            text = stripped

    return text


def _append_unique(values: list[str], value: str) -> None:
    if value and value not in values:
        values.append(value)


def split_part_numbers(
    value: Any,
    *,
    strip_rn_suffix: bool = True,
    remove_leading_zeroes_for_numeric_codes: bool = True,
) -> list[str]:
    """Split a cell with one or many part numbers into normalized keys."""
    if value is None or pd.isna(value):
        return []

    raw = str(value).strip()
    if not raw:
        return []

    result: list[str] = []
    parts = re.split(r"[,;/]+", raw)

    for part in parts:
        normalized = normalize_part(
            part,
            remove_leading_zeroes_for_numeric_codes=False,
        )
        if not normalized:
            continue

        _append_unique(result, normalized)

        if remove_leading_zeroes_for_numeric_codes and normalized.isdigit():
            stripped = normalized.lstrip("0")
            if stripped and stripped != normalized:
                _append_unique(result, stripped)

        if strip_rn_suffix and normalized.endswith("RN") and len(normalized) > 2:
            _append_unique(result, normalized[:-2])

    return result


def load_config(config_path: Path) -> dict[str, Any]:
    with config_path.open("r", encoding="utf-8") as config_file:
        config = json.load(config_file)

    config.setdefault("comparison", {})
    config["comparison"] = {**DEFAULT_COMPARISON, **config["comparison"]}
    config.setdefault("xml", {})
    config["xml"] = {**DEFAULT_XML_OPTIONS, **config["xml"]}
    return config


def detect_header_row(
    xlsx_path: Path,
    sheet_name: str | int | None,
    expected_columns: dict[str, str | None],
) -> int:
    raw = pd.read_excel(xlsx_path, sheet_name=sheet_name or 0, header=None, dtype=object)
    expected = {v for v in expected_columns.values() if v}

    best_index = 0
    best_score = -1
    for index, row in raw.iterrows():
        values = {str(value).strip() for value in row.dropna().tolist()}
        score = len(values & expected)
        if score > best_score:
            best_index = int(index)
            best_score = score

    if best_score <= 0:
        raise ValueError("Не удалось определить строку заголовков Excel.")

    return best_index


def _column_value(row: pd.Series, column_name: str | None, default: Any = "") -> Any:
    if not column_name or column_name not in row.index:
        return default
    value = row[column_name]
    if value is None or pd.isna(value):
        return default
    return value


def _to_number(value: Any) -> float | None:
    if value is None or pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip().replace("\u00a0", "").replace(" ", "")
    text = text.replace(",", ".")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def _clean_text(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def read_stock_excel(config: dict[str, Any]) -> list[StockItem]:
    xlsx_path = Path(config["stock_xlsx_path"])
    sheet_name = config.get("stock_sheet_name")
    header_row = config.get("stock_header_row")
    columns = config["columns"]
    comparison = config["comparison"]

    if header_row is None:
        header_index = detect_header_row(xlsx_path, sheet_name, columns)
    else:
        header_index = int(header_row) - 1

    df = pd.read_excel(
        xlsx_path,
        sheet_name=sheet_name or 0,
        header=header_index,
        dtype=object,
    )
    df = df.dropna(how="all")

    required = [columns["code_1c"], columns["name"], columns["qty"]]
    missing = [column for column in required if column and column not in df.columns]
    if missing:
        raise ValueError(f"В Excel не найдены обязательные колонки: {missing}")

    stock_items: list[StockItem] = []
    for idx, row in df.iterrows():
        code_1c = _clean_text(_column_value(row, columns.get("code_1c")))
        name = _clean_text(_column_value(row, columns.get("name")))
        qty = _to_number(_column_value(row, columns.get("qty")))

        if not code_1c or not name or qty is None or qty <= 0:
            continue

        article = _clean_text(_column_value(row, columns.get("article")))
        catalog_number_2 = _clean_text(_column_value(row, columns.get("catalog_number_2")))
        nomenclature_guid = _clean_text(_column_value(row, columns.get("nomenclature_guid")))

        keys = split_part_numbers(
            article,
            strip_rn_suffix=comparison["strip_rn_suffix"],
            remove_leading_zeroes_for_numeric_codes=comparison[
                "remove_leading_zeroes_for_numeric_codes"
            ],
        )
        if comparison["use_catalog_number_2"]:
            for key in split_part_numbers(
                catalog_number_2,
                strip_rn_suffix=comparison["strip_rn_suffix"],
                remove_leading_zeroes_for_numeric_codes=comparison[
                    "remove_leading_zeroes_for_numeric_codes"
                ],
            ):
                _append_unique(keys, key)

        if not keys and not nomenclature_guid:
            continue

        stock_items.append(
            StockItem(
                code_1c=code_1c,
                article=article,
                catalog_number_2=catalog_number_2,
                producer=_clean_text(_column_value(row, columns.get("producer"))),
                alpha_code=_clean_text(_column_value(row, columns.get("alpha_code"))),
                nomenclature_guid=nomenclature_guid,
                name=name,
                qty=qty,
                price=_column_value(row, columns.get("price"), None),
                keys=keys,
                row_numbers=[int(idx) + 2],
            )
        )

    if comparison["aggregate_stock_by_1c_code"]:
        return aggregate_stock(stock_items)

    return stock_items


def aggregate_stock(items: list[StockItem]) -> list[StockItem]:
    grouped: dict[str, StockItem] = {}

    for item in items:
        if item.code_1c not in grouped:
            grouped[item.code_1c] = item
            continue

        target = grouped[item.code_1c]
        target.qty += item.qty
        target.row_numbers.extend(item.row_numbers)
        for key in item.keys:
            _append_unique(target.keys, key)
        if not target.price and item.price:
            target.price = item.price

    return list(grouped.values())


def _child_text(node: Any, name: str) -> str:
    for child in node:
        if etree.QName(child).localname == name:
            return child.text.strip() if child.text else ""
    return ""


def _set_child_text(node: Any, name: str, value: Any) -> None:
    for child in node:
        if etree.QName(child).localname == name:
            child.text = str(value)
            return
    new_child = etree.SubElement(node, name)
    new_child.text = str(value)


def read_avito_xml(config: dict[str, Any]) -> tuple[Any, list[XmlAd]]:
    parser = etree.XMLParser(remove_blank_text=False, recover=True)
    tree = etree.parse(str(config["xml_path"]), parser)
    comparison = config["comparison"]

    ads: list[XmlAd] = []
    for index, node in enumerate(tree.xpath("//*[local-name()='Ad']")):
        oem = _child_text(node, "OEM")
        ads.append(
            XmlAd(
                index=index,
                ad_id=_child_text(node, "Id"),
                avito_id=_child_text(node, "AvitoId"),
                oem=oem,
                title=_child_text(node, "Title"),
                price=_child_text(node, "Price"),
                manager_name=_child_text(node, "ManagerName"),
                contact_phone=_child_text(node, "ContactPhone"),
                keys=split_part_numbers(
                    oem,
                    strip_rn_suffix=comparison["strip_rn_suffix"],
                    remove_leading_zeroes_for_numeric_codes=comparison[
                        "remove_leading_zeroes_for_numeric_codes"
                    ],
                ),
                node=node,
            )
        )

    return tree, ads


def stock_to_record(item: StockItem) -> dict[str, Any]:
    return {
        "Код 1С": item.code_1c,
        "Артикул": item.article,
        "Кат. номер2": item.catalog_number_2,
        "Производитель": item.producer,
        "Код альфа": item.alpha_code,
        "GUID": item.nomenclature_guid,
        "Номенклатура": item.name,
        "Остаток": item.qty,
        "Цена": item.price,
        "Ключи": ", ".join(item.keys),
        "Строки Excel": ", ".join(str(number) for number in item.row_numbers),
    }


def ad_to_record(ad: XmlAd) -> dict[str, Any]:
    return {
        "Id": ad.ad_id,
        "AvitoId": ad.avito_id,
        "OEM": ad.oem,
        "Title": ad.title,
        "Price": ad.price,
        "ManagerName": ad.manager_name,
        "ContactPhone": ad.contact_phone,
        "Ключи": ", ".join(ad.keys),
        "Порядок XML": ad.index + 1,
    }


def compare(stock_items: list[StockItem], xml_ads: list[XmlAd]) -> dict[str, list[dict[str, Any]]]:
    stock_by_key: dict[str, list[StockItem]] = defaultdict(list)
    xml_by_key: dict[str, list[XmlAd]] = defaultdict(list)

    for item in stock_items:
        for key in item.keys:
            stock_by_key[key].append(item)

    for ad in xml_ads:
        for key in ad.keys:
            xml_by_key[key].append(ad)

    xml_duplicates: list[dict[str, Any]] = []
    for key, ads in sorted(xml_by_key.items()):
        unique_ids = sorted({ad.ad_id for ad in ads})
        if len(unique_ids) > 1:
            xml_duplicates.append(
                {
                    "Ключ": key,
                    "Количество объявлений": len(unique_ids),
                    "Id XML": ", ".join(unique_ids),
                    "Title": " | ".join(ad.title for ad in ads[:5]),
                }
            )

    matched_stock_codes: set[str] = set()
    matched_ad_ids: set[str] = set()
    report_match: list[dict[str, Any]] = []
    report_remove: list[dict[str, Any]] = []
    report_manual: list[dict[str, Any]] = []

    for ad in xml_ads:
        if not ad.keys:
            row = ad_to_record(ad)
            row["Причина"] = "В XML не заполнен OEM или не удалось получить ключ сравнения"
            report_manual.append(row)
            continue

        candidate_items: dict[str, StockItem] = {}
        matched_keys: list[str] = []
        for key in ad.keys:
            for item in stock_by_key.get(key, []):
                candidate_items[item.code_1c] = item
                _append_unique(matched_keys, key)

        if not candidate_items:
            row = ad_to_record(ad)
            row["Причина"] = "Нет совпадения в остатках 1С"
            report_remove.append(row)
            continue

        if len(candidate_items) > 1:
            row = ad_to_record(ad)
            row["Причина"] = "Один OEM XML совпал с несколькими товарами 1С"
            row["Коды 1С"] = ", ".join(candidate_items.keys())
            report_manual.append(row)
            continue

        item = next(iter(candidate_items.values()))
        matched_stock_codes.add(item.code_1c)
        matched_ad_ids.add(ad.ad_id)
        row = {**ad_to_record(ad), **stock_to_record(item)}
        row["Совпавшие ключи"] = ", ".join(matched_keys)
        report_match.append(row)

    report_add: list[dict[str, Any]] = []
    for item in stock_items:
        if not item.keys:
            row = stock_to_record(item)
            row["Причина"] = "В строке остатков нет артикула или каталожного номера"
            report_manual.append(row)
            continue

        if item.code_1c in matched_stock_codes:
            continue

        has_xml_match = any(xml_by_key.get(key) for key in item.keys)
        if not has_xml_match:
            report_add.append(stock_to_record(item))

    return {
        "add": report_add,
        "remove": report_remove,
        "match": report_match,
        "manual": report_manual,
        "xml_duplicates": xml_duplicates,
    }


def apply_xml_changes(tree: Any, result: dict[str, list[dict[str, Any]]], output_xml: Path, update_price: bool) -> None:
    remove_ids = {row["Id"] for row in result["remove"]}
    price_by_ad_id = {
        row["Id"]: row.get("Цена")
        for row in result["match"]
        if row.get("Id") and row.get("Цена") not in (None, "")
    }

    for ad in list(tree.xpath("//*[local-name()='Ad']")):
        ad_id = _child_text(ad, "Id")
        if ad_id in remove_ids:
            parent = ad.getparent()
            if parent is not None:
                parent.remove(ad)
            continue

        if update_price and ad_id in price_by_ad_id:
            _set_child_text(ad, "Price", price_by_ad_id[ad_id])

    output_xml.parent.mkdir(parents=True, exist_ok=True)
    tree.write(
        str(output_xml),
        encoding="utf-8",
        xml_declaration=True,
        pretty_print=False,
    )


def write_report(
    output_path: Path,
    result: dict[str, list[dict[str, Any]]],
    summary: dict[str, Any],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        pd.DataFrame(
            [{"Показатель": key, "Значение": value} for key, value in summary.items()]
        ).to_excel(writer, sheet_name="Итог", index=False)
        pd.DataFrame(result["add"]).to_excel(writer, sheet_name="Добавить на Авито", index=False)
        pd.DataFrame(result["remove"]).to_excel(writer, sheet_name="Снять с Авито", index=False)
        pd.DataFrame(result["match"]).to_excel(writer, sheet_name="Совпадают", index=False)
        pd.DataFrame(result["manual"]).to_excel(writer, sheet_name="Ручная проверка", index=False)
        pd.DataFrame(result["xml_duplicates"]).to_excel(writer, sheet_name="Дубли XML по OEM", index=False)


def run(config: dict[str, Any]) -> dict[str, Any]:
    output_dir = Path(config["output_dir"])
    output_xml = output_dir / "avito_autoload_filtered.xml"
    output_report = output_dir / "compare_report.xlsx"

    stock_items = read_stock_excel(config)
    tree, xml_ads = read_avito_xml(config)
    result = compare(stock_items, xml_ads)

    xml_keys = sorted({key for ad in xml_ads for key in ad.keys})
    summary = {
        "строк в Excel после фильтрации": len(stock_items),
        "уникальных SKU 1С": len({item.code_1c for item in stock_items}),
        "объявлений в исходном XML": len(xml_ads),
        "уникальных OEM/ключей в XML": len(xml_keys),
        "оставлено в XML": len(xml_ads) - len(result["remove"]),
        "добавить на Авито": len(result["add"]),
        "снять с Авито": len(result["remove"]),
        "совпадают": len(result["match"]),
        "ручная проверка": len(result["manual"]),
        "дубли XML по OEM": len(result["xml_duplicates"]),
        "путь к итоговому XML": str(output_xml),
    }

    apply_xml_changes(
        tree,
        result,
        output_xml,
        update_price=bool(config["xml"]["update_price"]),
    )
    write_report(output_report, result, summary)

    return {
        "summary": summary,
        "output_xml": output_xml,
        "output_report": output_report,
        "result": result,
    }


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Сравнить XML Авито с Excel-остатками 1С и сформировать новый XML."
    )
    parser.add_argument("--config", required=True, help="Путь к JSON-конфигурации.")
    args = parser.parse_args()

    config = load_config(Path(args.config))
    result = run(config)

    summary = result["summary"]
    print("Готово.")
    print(f"Исходных объявлений XML: {summary['объявлений в исходном XML']}")
    print(f"Строк остатков после фильтрации: {summary['строк в Excel после фильтрации']}")
    print(f"Уникальных SKU 1С: {summary['уникальных SKU 1С']}")
    print(f"Оставлено в XML: {summary['оставлено в XML']}")
    print(f"Снято с Авито: {summary['снять с Авито']}")
    print(f"Добавить на Авито: {summary['добавить на Авито']}")
    print(f"Ручная проверка: {summary['ручная проверка']}")
    print(f"Дубли XML по OEM: {summary['дубли XML по OEM']}")
    print(f"Итоговый XML: {result['output_xml']}")
    print(f"Отчёт: {result['output_report']}")


if __name__ == "__main__":
    main()
