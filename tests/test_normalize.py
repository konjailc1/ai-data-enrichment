"""Rule tests for price, phone, address, and duplicate detection."""

from __future__ import annotations

import unittest

from enrich import (
    SourceRow,
    address_is_valid,
    decide_status,
    failed_row,
    mark_duplicates,
    normalize_phone,
    parse_price_vnd,
    polish_address,
    rule_address,
)


def source(**overrides: object) -> SourceRow:
    base = dict(
        row_id=1,
        raw_product="sua",
        raw_price="25k",
        customer_name="An",
        raw_phone="0912345678",
        raw_address="45 ngoc lam",
        raw_feedback="ngon",
        price_vnd=25000,
        phone="0912345678",
        product_empty=False,
    )
    base.update(overrides)
    return SourceRow(**base)  # type: ignore[arg-type]


class PriceTests(unittest.TestCase):
    def test_formats_from_the_brief(self) -> None:
        self.assertEqual(parse_price_vnd("25.000đ"), 25000)
        self.assertEqual(parse_price_vnd("25k"), 25000)
        self.assertEqual(parse_price_vnd("25,000"), 25000)
        self.assertEqual(parse_price_vnd("1tr2"), 1_200_000)

    def test_more_vietnamese_forms(self) -> None:
        self.assertEqual(parse_price_vnd("1.200.000"), 1_200_000)
        self.assertEqual(parse_price_vnd("12.5k"), 12500)
        self.assertEqual(parse_price_vnd("1,2tr"), 1_200_000)
        self.assertEqual(parse_price_vnd("2 triệu"), 2_000_000)
        self.assertEqual(parse_price_vnd("1tr rưỡi"), 1_500_000)
        self.assertEqual(parse_price_vnd("4.500"), 4500)
        self.assertEqual(parse_price_vnd("  89.000 đ "), 89000)
        self.assertEqual(parse_price_vnd("9500"), 9500)

    def test_rejects_unparseable(self) -> None:
        self.assertIsNone(parse_price_vnd("lien he"))
        self.assertIsNone(parse_price_vnd(""))
        self.assertIsNone(parse_price_vnd("abc"))
        self.assertIsNone(parse_price_vnd("25"))
        self.assertIsNone(parse_price_vnd(None))


class PhoneTests(unittest.TestCase):
    def test_formats_from_the_brief(self) -> None:
        self.assertEqual(normalize_phone("0912 345 678"), "0912345678")
        self.assertEqual(normalize_phone("+84912345678"), "0912345678")
        self.assertEqual(normalize_phone("912.345.678"), "0912345678")

    def test_dashes_and_invalid(self) -> None:
        self.assertEqual(normalize_phone("0988-111-222"), "0988111222")
        self.assertIsNone(normalize_phone("0123 456 789"))
        self.assertIsNone(normalize_phone("12345"))
        self.assertIsNone(normalize_phone(""))
        self.assertIsNone(normalize_phone(None))


class AddressTests(unittest.TestCase):
    def test_2025_form(self) -> None:
        good = "45 Ngọc Lâm, Phường Bồ Đề, TP. Hà Nội"
        self.assertTrue(address_is_valid(good))
        self.assertEqual(polish_address("45 Ngọc Lâm, phường Bồ Đề, tp. hà nội"), good)

    def test_gazetteer_overrides_misspellings(self) -> None:
        self.assertEqual(
            rule_address("8 duong kieu ky, xa kieu ky, huyen gia lam"),
            "8 Kiêu Kỵ, Xã Gia Lâm, TP. Hà Nội",
        )
        self.assertEqual(
            rule_address("4 dong la, an khanh"),
            "4 Đông La, Xã An Khánh, TP. Hà Nội",
        )
        self.assertEqual(
            rule_address("15 tay mo, nam tu liem"),
            "15 Tây Mỗ, Phường Tây Mỗ, TP. Hà Nội",
        )
        self.assertEqual(
            rule_address("27 co linh long bien"),
            "27 Cổ Linh, Phường Long Biên, TP. Hà Nội",
        )
        self.assertIsNone(rule_address("gan cho, hn"))
        self.assertIsNone(rule_address("co linh khong ro so nha"))

    def test_rejects_district(self) -> None:
        self.assertFalse(
            address_is_valid("45 Ngọc Lâm, Phường Ngọc Lâm, Quận Long Biên, TP. Hà Nội")
        )
        self.assertFalse(address_is_valid(""))


class StatusTests(unittest.TestCase):
    def test_duplicate_is_flagged_once(self) -> None:
        first = source(row_id=1)
        second = source(row_id=2)
        mark_duplicates([first, second])
        self.assertFalse(first.is_duplicate)
        self.assertTrue(second.is_duplicate)

    def test_review_reasons(self) -> None:
        row = failed_row(
            source(
                raw_product="",
                product_empty=True,
                raw_price="lien he",
                price_vnd=None,
                raw_phone="12345",
                phone=None,
                is_duplicate=True,
            )
        )
        decide_status(row)
        self.assertIn("NEEDS REVIEW", row.status)
        self.assertTrue(any("Empty product" in reason for reason in row.reasons))
        self.assertTrue(any("Unparseable price" in reason for reason in row.reasons))
        self.assertTrue(any("Invalid phone" in reason for reason in row.reasons))
        self.assertTrue(any("Duplicate" in reason for reason in row.reasons))


if __name__ == "__main__":
    unittest.main()
