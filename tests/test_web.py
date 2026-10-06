"""The web page, driven through Flask's test client with jobs run synchronously."""
import io
import json
import os

import openpyxl
import pytest

from course_review.web import SyncExecutor, create_app

from helpers import add, checklist, new_doc, zip_dir
from test_fallback import package_with_roman_pop_quiz


@pytest.fixture
def app(tmp_path):
    a = create_app(jobs_dir=str(tmp_path / "jobs"), checklist=checklist(tmp_path), books_dir=str(tmp_path / "books"), executor=SyncExecutor())
    a.config["TESTING"] = True
    return a


def zip_bytes(tmp_path):
    return open(package_with_roman_pop_quiz(tmp_path), "rb").read()


def upload(client, data, **fields):
    form = {"package": (io.BytesIO(data), "course.zip"), **fields}
    return client.post("/jobs", data=form, content_type="multipart/form-data")


def test_home_page_has_the_upload_form(app):
    r = app.test_client().get("/")
    assert r.status_code == 200 and b'type="file" name="package"' in r.data and b"Start review" in r.data
    assert b"default-src" in r.headers["Content-Security-Policy"].encode()


def test_upload_runs_a_review_and_serves_the_results(app, tmp_path):
    c = app.test_client()
    r = upload(c, zip_bytes(tmp_path))
    assert r.status_code == 302
    job = r.headers["Location"].rsplit("/", 1)[-1]
    status = c.get(f"/jobs/{job}/status").get_json()
    assert status["state"] == "done" and "documents" in status["summary"]
    page = c.get(f"/jobs/{job}")
    assert b"Filled workbook" in page.data and b"course.zip" in page.data
    xlsx = c.get(f"/jobs/{job}/files/Course-Review-Checklist-filled.xlsx")
    assert xlsx.status_code == 200 and "attachment" in xlsx.headers["Content-Disposition"]
    wb = openpyxl.load_workbook(io.BytesIO(xlsx.data))
    assert "Review Summary" in wb.sheetnames
    rep = c.get(f"/jobs/{job}/files/report.html")
    assert rep.status_code == 200 and b"default-src 'none'" in rep.headers["Content-Security-Policy"].encode()
    assert json.loads(c.get(f"/jobs/{job}/files/review.json").data)["findings"]
    z = c.get(f"/jobs/{job}/files/Marked-up-documents.zip")
    assert z.status_code == 200 and "attachment" in z.headers["Content-Disposition"] and z.data[:2] == b"PK"
    assert b"Download fixed documents" in page.data


def test_downloads_work_with_relative_folders_as_serve_uses_them(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    a = create_app(jobs_dir="web_jobs", checklist=checklist(tmp_path), books_dir="book_indexes", executor=SyncExecutor())
    c = a.test_client()
    job = upload(c, zip_bytes(tmp_path)).headers["Location"].rsplit("/", 1)[-1]
    for name in ("Course-Review-Checklist-filled.xlsx", "report.html", "review.json", "Marked-up-documents.zip"):
        assert c.get(f"/jobs/{job}/files/{name}").status_code == 200, name


def test_the_uploaded_zip_is_not_kept(app, tmp_path):
    c = app.test_client()
    job = upload(c, zip_bytes(tmp_path)).headers["Location"].rsplit("/", 1)[-1]
    files = os.listdir(tmp_path / "jobs" / job)
    assert "package.zip" not in files and "report.html" in files and "meta.json" in files


def test_rejects_a_missing_or_non_zip_upload(app):
    c = app.test_client()
    assert c.post("/jobs", data={}, content_type="multipart/form-data").status_code == 400
    r = c.post("/jobs", data={"package": (io.BytesIO(b"x"), "notes.txt")}, content_type="multipart/form-data")
    assert r.status_code == 400 and b"must be a .zip" in r.data


def test_a_corrupt_zip_fails_the_job_with_a_message_not_the_server(app):
    c = app.test_client()
    job = upload(c, b"this is not a zip").headers["Location"].rsplit("/", 1)[-1]
    s = c.get(f"/jobs/{job}/status").get_json()
    assert s["state"] == "failed" and s["error"]
    assert b"could not finish" in c.get(f"/jobs/{job}").data


def test_unknown_jobs_and_unsafe_file_names_are_404(app, tmp_path):
    c = app.test_client()
    assert c.get("/jobs/000000000000").status_code == 404
    assert c.get("/jobs/..%2F..%2Fetc/status").status_code == 404
    job = upload(c, zip_bytes(tmp_path)).headers["Location"].rsplit("/", 1)[-1]
    for name in ("meta.json", "..%2Fmeta.json", "package.zip", "checklist.xlsx"):
        assert c.get(f"/jobs/{job}/files/{name}").status_code == 404


def test_model_settings_that_cannot_be_reached_fail_the_job_clearly(app, tmp_path):
    c = app.test_client()
    job = upload(c, zip_bytes(tmp_path), model_mode="fallback", model_url="http://127.0.0.1:1", model_name="m").headers["Location"].rsplit("/", 1)[-1]
    s = c.get(f"/jobs/{job}/status").get_json()
    assert s["state"] == "failed" and "not usable" in s["error"]


def test_a_review_with_its_own_checklist_upload(app, tmp_path):
    c = app.test_client()
    (tmp_path / "x").mkdir()
    own = open(checklist(tmp_path / "x"), "rb").read()
    job = upload(c, zip_bytes(tmp_path), checklist=(io.BytesIO(own), "mine.xlsx")).headers["Location"].rsplit("/", 1)[-1]
    assert c.get(f"/jobs/{job}/status").get_json()["state"] == "done"
    bad = c.post("/jobs", data={"package": (io.BytesIO(b"x"), "a.zip"), "checklist": (io.BytesIO(b"x"), "mine.txt")}, content_type="multipart/form-data")
    assert bad.status_code == 400


def test_delete_removes_the_job_and_listing(app, tmp_path):
    c = app.test_client()
    job = upload(c, zip_bytes(tmp_path)).headers["Location"].rsplit("/", 1)[-1]
    assert b"course.zip" in c.get("/").data
    assert c.post(f"/jobs/{job}/delete").status_code == 302
    assert c.get(f"/jobs/{job}").status_code == 404 and not (tmp_path / "jobs" / job).exists()


def test_a_job_interrupted_by_a_restart_is_marked_failed(tmp_path):
    jobs = tmp_path / "jobs"
    d = jobs / "abcdef012345"
    d.mkdir(parents=True)
    (d / "meta.json").write_text(json.dumps({"id": "abcdef012345", "name": "x.zip", "state": "running"}), encoding="utf8")
    a = create_app(jobs_dir=str(jobs), checklist=None, executor=SyncExecutor())
    s = a.test_client().get("/jobs/abcdef012345/status").get_json()
    assert s["state"] == "failed" and "restarted" in s["error"]


def test_textbook_indexes_in_the_books_folder_are_offered(tmp_path):
    (tmp_path / "books" / "islamiat" / "pages").mkdir(parents=True)
    a = create_app(jobs_dir=str(tmp_path / "jobs"), checklist=None, books_dir=str(tmp_path / "books"), executor=SyncExecutor())
    assert b"<option>islamiat</option>" in a.test_client().get("/").data


def test_status_file_survives_a_locked_replace_and_progress_errors_do_not_stop_a_review(tmp_path, monkeypatch):
    # Windows refuses os.replace while another handle has the file open; the write must retry, not fail.
    import course_review.web as W
    real, calls = W.os.replace, {"n": 0}

    def flaky(a, b):
        calls["n"] += 1
        if calls["n"] <= 3:
            raise PermissionError(5, "Access is denied")
        return real(a, b)
    monkeypatch.setattr(W.os, "replace", flaky)
    W._write_meta(tmp_path, {"state": "running"})
    assert W._read_meta(tmp_path) == {"state": "running"} and calls["n"] == 4
