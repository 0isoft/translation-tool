import unittest

from starlette.requests import Request

from app.main import request_origin


def make_request(headers: dict[str, str]) -> Request:
    return Request({
        "type": "http",
        "method": "GET",
        "scheme": "http",
        "path": "/manifest.xml",
        "raw_path": b"/manifest.xml",
        "query_string": b"",
        "server": ("internal", 8000),
        "headers": [
            (name.lower().encode(), value.encode())
            for name, value in headers.items()
        ],
    })


class ManifestOriginTests(unittest.TestCase):
    def test_vercel_forwarded_origin_is_used(self):
        request = make_request({
            "host": "internal",
            "x-forwarded-proto": "https",
            "x-forwarded-host": "translation.example.com",
        })

        self.assertEqual(
            request_origin(request),
            "https://translation.example.com",
        )

    def test_request_origin_is_used_without_proxy_headers(self):
        request = make_request({"host": "localhost:8000"})

        self.assertEqual(request_origin(request), "http://localhost:8000")


if __name__ == "__main__":
    unittest.main()
