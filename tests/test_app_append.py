from pathlib import Path

from streamlit.testing.v1 import AppTest


def test_streamlit_append_mode_processes_uploaded_base_and_new_file():
    app_path = Path(__file__).resolve().parents[1] / "app.py"
    app = AppTest.from_file(str(app_path), default_timeout=30).run()
    app.radio[0].set_value("Append").run(timeout=30)

    new_upload = app.file_uploader[0]
    base_upload = next(uploader for uploader in app.file_uploader if uploader.key == "append_base_upload")
    new_upload.set_value(("new.csv", b"Order ID,Nama\nA2,Andi\n", "text/csv"))
    base_upload.set_value(("base.csv", b"Order ID,Nama\nA1,Budi\n", "text/csv"))
    app.run(timeout=30)

    assert not app.exception
    metrics = {metric.label: metric.value for metric in app.metric}
    assert metrics["Baris Base"] == "1"
    assert metrics["Ditambahkan"] == "1"
    assert metrics["Dilewati"] == "0"
    assert metrics["Baris Akhir"] == "2"
    assert any("Rincian per file" in item.value for item in app.markdown)