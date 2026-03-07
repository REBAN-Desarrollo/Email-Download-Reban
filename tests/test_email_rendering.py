import unittest


from email_rendering import (
    build_gmail_print_document,
    build_original_email_document,
    sanitize_email_html,
)


class EmailRenderingTests(unittest.TestCase):
    def test_sanitize_email_html_removes_document_shell_and_unsafe_layout_rules(self):
        source_html = """
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { margin-top: 480px; }
  </style>
  <script>alert("x")</script>
</head>
<body style="margin-top:300px; position:absolute; top:25px; color:#112233; font-weight:bold;">
  <!--[if gte mso 9]><xml><x:OfficeDocumentSettings/></xml><![endif]-->
  <table style="position:absolute; left:30px; margin-top:100px;">
    <tr>
      <td><img src="data:image/png;base64,abc123" style="width:120px; position:absolute; top:4px;"></td>
      <td><span style="color:#224466; font-weight:bold;">Hola mundo</span></td>
    </tr>
  </table>
</body>
</html>
"""

        sanitized = sanitize_email_html(source_html)

        self.assertNotIn("<html", sanitized.lower())
        self.assertNotIn("<head", sanitized.lower())
        self.assertNotIn("<body", sanitized.lower())
        self.assertNotIn("<script", sanitized.lower())
        self.assertNotIn("position:absolute", sanitized.lower())
        self.assertNotIn("margin-top", sanitized.lower())
        self.assertIn("Hola mundo", sanitized)
        self.assertIn("data:image/png;base64,abc123", sanitized)
        self.assertIn("color:#224466", sanitized.replace(" ", ""))

    def test_build_gmail_print_document_wraps_fragment_in_gmail_shell(self):
        document = build_gmail_print_document(
            account_email="danielgarza@reban.com",
            subject="FW: Propuesta",
            sender='"Juan Pablo Castellanos" <comercial@reban.com>',
            recipient="<danielgarza@reban.com>",
            cc="",
            sent_at="2017-12-04 10:42",
            body_fragment="<p>Contenido</p>",
        )

        self.assertIn("Gmail", document)
        self.assertIn("1 mensaje", document)
        self.assertIn("FW: Propuesta", document)
        self.assertIn("danielgarza@reban.com", document)
        self.assertIn("<p>Contenido</p>", document)

    def test_build_original_email_document_wraps_fragment_as_standalone_html(self):
        document = build_original_email_document("<div>Original</div>")

        self.assertIn("<html", document.lower())
        self.assertIn("<meta charset=\"utf-8\">", document.lower())
        self.assertIn("<div>Original</div>", document)


if __name__ == "__main__":
    unittest.main()
