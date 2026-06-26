from shopee_scraper_core import extract_item_id, extract_shop_id_from_link, parse_proxy

def test_extract_item_id():
    assert extract_item_id("54005267037") == "54005267037"
    assert extract_item_id("https://shopee.vn/product/302865535/54005267037") == "54005267037"
    assert extract_item_id("https://shopee.vn/Ao-i.302865535.54005267037") == "54005267037"
    assert extract_item_id("invalid-link") is None

def test_extract_shop_id_from_link():
    assert extract_shop_id_from_link("https://shopee.vn/product/302865535/54005267037") == "302865535"
    assert extract_shop_id_from_link("https://shopee.vn/Ao-i.302865535.54005267037") == "302865535"
    assert extract_shop_id_from_link("invalid-link") is None

def test_parse_proxy():
    assert parse_proxy("127.0.0.1:8080") == {"http": "http://127.0.0.1:8080", "https": "http://127.0.0.1:8080"}
    assert parse_proxy("127.0.0.1:8080:user:pass") == {"http": "http://user:pass@127.0.0.1:8080", "https": "http://user:pass@127.0.0.1:8080"}
    assert parse_proxy("") is None
