import contextlib
import io
import importlib.util
from pathlib import Path
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location("fanjian", ROOT / "fanjian.py")
f = importlib.util.module_from_spec(spec)
spec.loader.exec_module(f)

class ConversionTest(unittest.TestCase):
    def test_sample(self):
        table, lens, tw, user = f.build()
        ctx = table, lens, tw, user, f.load_at(f.LEX / "post.txt")
        with tempfile.TemporaryDirectory() as d:
            path = Path(d) / "sample.txt"
            path.write_text((ROOT / "tests/sample.txt").read_text(encoding="utf-8"), encoding="utf-8")
            with contextlib.redirect_stdout(io.StringIO()):
                f.convert(path, ctx)
            out = (Path(d) / "sample-b5.txt").read_bytes().decode("cp950")
        self.assertIn("第1章　出發", out.replace("第一章", "第1章"))
        self.assertIn("面前", out)
        self.assertIn("丹田裡", out)
        self.assertIn("數千里", out)
        self.assertIn("「你好『朋友』」？！～ ●", out)
        self.assertIn("~\\\r\n", out)

if __name__ == "__main__":
    unittest.main()
