import json
from pathlib import Path

import pandas as pd
from lxml import etree

from prototype.compare_avito_xml_with_1c_stock import load_config, run


def test_compare_small_dataset(tmp_path: Path):
    xml_path = tmp_path / "avito.xml"
    stock_path = tmp_path / "stock.xlsx"
    output_dir = tmp_path / "output"
    config_path = tmp_path / "config.json"

    xml_path.write_text(
        """<?xml version="1.0" encoding="UTF-8"?>
<Ads target="Avito.ru" formatVersion="3">
  <Ad>
    <Id>ad-1964793</Id>
    <Title>Item 1964793</Title>
    <Price>100</Price>
    <OEM>1964793</OEM>
    <Images>
      <Image url="https://site.ru/1964793.jpg"/>
    </Images>
  </Ad>
  <Ad>
    <Id>ad-2705323</Id>
    <Title>Item 2705323</Title>
    <Price>200</Price>
    <OEM>2705323</OEM>
  </Ad>
  <Ad>
    <Id>ad-3918776</Id>
    <Title>Item 3918776</Title>
    <Price>300</Price>
    <OEM>3918776</OEM>
  </Ad>
</Ads>
""",
        encoding="utf-8",
    )

    pd.DataFrame(
        [
            {
                "Номенклатура.Код": "RS-1",
                "Артикул": "1964793",
                "Кат. номер2": "",
                "Идентификатор номенклатуры": "guid-1",
                "Номенклатура": "Stock 1964793",
                "Остаток": 1,
            },
            {
                "Номенклатура.Код": "RS-2",
                "Артикул": "3918776RN",
                "Кат. номер2": "",
                "Идентификатор номенклатуры": "guid-2",
                "Номенклатура": "Stock 3918776RN",
                "Остаток": 1,
            },
            {
                "Номенклатура.Код": "RS-3",
                "Артикул": "9999999",
                "Кат. номер2": "",
                "Идентификатор номенклатуры": "guid-3",
                "Номенклатура": "Stock 9999999",
                "Остаток": 1,
            },
        ]
    ).to_excel(stock_path, index=False)

    config_path.write_text(
        json.dumps(
            {
                "xml_path": str(xml_path),
                "stock_xlsx_path": str(stock_path),
                "output_dir": str(output_dir),
                "stock_sheet_name": None,
                "stock_header_row": None,
                "columns": {
                    "code_1c": "Номенклатура.Код",
                    "article": "Артикул",
                    "producer": None,
                    "alpha_code": None,
                    "catalog_number_2": "Кат. номер2",
                    "nomenclature_guid": "Идентификатор номенклатуры",
                    "name": "Номенклатура",
                    "qty": "Остаток",
                    "price": None,
                },
                "comparison": {
                    "strip_rn_suffix": True,
                    "remove_leading_zeroes_for_numeric_codes": True,
                    "use_catalog_number_2": True,
                    "aggregate_stock_by_1c_code": True,
                },
                "xml": {
                    "update_price": False,
                    "preserve_order": True,
                },
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    result = run(load_config(config_path))

    matched_ids = {row["Id"] for row in result["result"]["match"]}
    removed_ids = {row["Id"] for row in result["result"]["remove"]}
    add_articles = {row["Артикул"] for row in result["result"]["add"]}

    assert "ad-1964793" in matched_ids
    assert "ad-3918776" in matched_ids
    assert "ad-2705323" in removed_ids
    assert "9999999" in add_articles

    output_xml = output_dir / "avito_autoload_filtered.xml"
    report = output_dir / "compare_report.xlsx"
    assert output_xml.exists()
    assert report.exists()

    tree = etree.parse(str(output_xml))
    remaining_ids = [node.text for node in tree.xpath("//*[local-name()='Id']")]
    assert remaining_ids == ["ad-1964793", "ad-3918776"]

    workbook = pd.ExcelFile(report)
    assert set(workbook.sheet_names) == {
        "Итог",
        "Добавить на Авито",
        "Снять с Авито",
        "Совпадают",
        "Ручная проверка",
        "Дубли XML по OEM",
    }
