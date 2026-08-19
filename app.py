import json
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_from_directory

from services.cls_service import CLSService
from services.enrollment_service import EnrollmentService
from services.services_service import ServicesService
from services.structure_service import StructureService
from services.timeliness_service import TimelinessService


BASE_DIR = Path(__file__).resolve().parent


def create_app(test_config=None):
    app = Flask(__name__)
    app.config.update(
        SECRET_KEY="local-weekly-reporting-app",
        DATA_DIR=BASE_DIR / "data",
        OUTPUT_DIR=BASE_DIR / "outputs",
        DEFAULT_SERVICE_USER_MAPPING_DIR=BASE_DIR / "sample_data",
        LOAD_DEFAULT_SERVICE_USER_MAPPINGS=None,
        MAX_CONTENT_LENGTH=25 * 1024 * 1024,
    )
    if test_config:
        app.config.update(test_config)

    load_default_mappings = app.config.get("LOAD_DEFAULT_SERVICE_USER_MAPPINGS")
    if load_default_mappings is None:
        load_default_mappings = not app.config.get("TESTING")
    default_mapping_dir = None
    if load_default_mappings:
        default_mapping_dir = Path(app.config["DEFAULT_SERVICE_USER_MAPPING_DIR"])
    service = StructureService(Path(app.config["DATA_DIR"]), default_mapping_dir)
    cls_service = CLSService(Path(app.config["OUTPUT_DIR"]))
    services_service = ServicesService(Path(app.config["OUTPUT_DIR"]))
    enrollment_service = EnrollmentService(Path(app.config["OUTPUT_DIR"]))
    timeliness_service = TimelinessService(Path(app.config["OUTPUT_DIR"]))

    @app.after_request
    def prevent_api_caching(response):
        if request.path.startswith("/api/"):
            response.headers["Cache-Control"] = "no-store"
        return response

    @app.get("/")
    def index():
        return render_template("index.html")

    @app.get("/api/programs/<program>/structure")
    def get_structure(program):
        return jsonify(service.load_structure(program))

    @app.put("/api/programs/<program>/structure")
    def save_structure(program):
        payload = request.get_json(silent=True) or {}
        try:
            saved = service.save_structure(program, payload)
        except ValueError as exc:
            return jsonify({"error": str(exc)}), 400
        return jsonify(saved)

    @app.get("/api/programs/<program>/metrics")
    def get_metrics(program):
        return jsonify(service.load_metrics(program))

    @app.get("/api/programs/<program>/service-users")
    def get_service_users(program):
        try:
            mappings = service.load_service_users(program)
        except ValueError as exc:
            return jsonify({"error": str(exc)}), 400
        return jsonify(mappings)

    @app.put("/api/programs/<program>/metrics")
    def save_metrics(program):
        payload = request.get_json(silent=True) or {}
        try:
            saved = service.save_metrics(program, payload)
        except ValueError as exc:
            return jsonify({"error": str(exc)}), 400
        return jsonify(saved)

    @app.post("/api/programs/<program>/mock")
    def load_mock(program):
        try:
            result = service.install_mock_data(program)
        except ValueError as exc:
            return jsonify({"error": str(exc)}), 400
        return jsonify(result)

    @app.get("/api/programs/<program>/presets")
    def list_presets(program):
        try:
            presets = service.list_presets(program)
        except ValueError as exc:
            return jsonify({"error": str(exc)}), 400
        return jsonify({"presets": presets})

    @app.post("/api/programs/<program>/presets")
    def save_preset(program):
        payload = request.get_json(silent=True) or {}
        try:
            result = service.save_preset(program, payload.get("name", ""), payload.get("structure", {}))
        except ValueError as exc:
            return jsonify({"error": str(exc)}), 400
        return jsonify(result)

    @app.post("/api/programs/<program>/presets/load")
    def load_preset(program):
        payload = request.get_json(silent=True) or {}
        try:
            result = service.load_preset(program, payload.get("name", ""))
        except ValueError as exc:
            return jsonify({"error": str(exc)}), 400
        return jsonify(result)

    @app.post("/api/programs/<program>/team/import")
    def import_team(program):
        uploaded = request.files.get("team_file")
        if not uploaded or not uploaded.filename:
            return jsonify({"error": "Drop a Team CSV or Excel file first."}), 400
        try:
            result = cls_service.structure_from_team_upload(program, uploaded)
            result["structure"] = service.save_structure(program, result["structure"])
            result["service_users"] = service.merge_service_users(
                program,
                result.get("staff_to_service_user", {}),
            )
        except ValueError as exc:
            return jsonify({"error": str(exc)}), 400
        result.pop("staff_to_manager", None)
        result.pop("service_user_to_staff", None)
        result.pop("staff_to_service_user", None)
        return jsonify(result)

    @app.post("/api/programs/<program>/team/export")
    def export_team(program):
        payload = request.get_json(silent=True) or {}
        try:
            structure = service.save_structure(program, payload)
            service_users = service.load_service_users(program)
            filename = service.export_team_mapping(
                program,
                structure,
                service_users,
                Path(app.config["OUTPUT_DIR"]),
            )
        except ValueError as exc:
            return jsonify({"error": str(exc)}), 400
        return jsonify({
            "structure": structure,
            "service_users": service_users,
            "download_name": filename,
            "download_url": f"/downloads/{program.upper()}/{filename}",
        })

    @app.post("/api/programs/<program>/cls/process")
    def process_cls(program):
        uploaded = request.files.get("cls_file")
        if not uploaded or not uploaded.filename:
            return jsonify({"error": "Drop a CLS CSV or Excel file first."}), 400
        try:
            result = cls_service.process(
                program=program,
                uploaded_file=uploaded,
                team_uploaded_file=request.files.get("team_file"),
                start_date=request.form.get("start_date", ""),
                end_date=request.form.get("end_date", ""),
                structure=service.load_structure(program),
                current_metrics=service.load_metrics(program),
            )
            service.save_structure(program, result["structure"])
            service.save_metrics(program, result["metrics"])
        except ValueError as exc:
            return jsonify({"error": str(exc)}), 400
        result["download_url"] = f"/downloads/{program.upper()}/{result['download_name']}"
        return jsonify(result)

    @app.post("/api/programs/<program>/services/process")
    def process_services(program):
        uploaded = request.files.get("services_file")
        if not uploaded or not uploaded.filename:
            return jsonify({"error": "Drop a Services CSV or Excel file first."}), 400
        try:
            result = services_service.process(
                program=program,
                uploaded_file=uploaded,
                start_date=request.form.get("start_date", ""),
                end_date=request.form.get("end_date", ""),
                structure=service.load_structure(program),
                current_metrics=service.load_metrics(program),
                service_users=service.load_service_users(program),
                service_user_teams=service.load_service_user_teams(program),
            )
            service.save_structure(program, result["structure"])
            service.save_metrics(program, result["metrics"])
        except ValueError as exc:
            return jsonify({"error": str(exc)}), 400
        result["download_url"] = f"/downloads/{program.upper()}/{result['download_name']}"
        return jsonify(result)

    @app.post("/api/programs/<program>/enrollment/process")
    def process_enrollment(program):
        uploaded = request.files.get("enrollment_file")
        if not uploaded or not uploaded.filename:
            return jsonify({"error": "Drop an Enrollment/Outcomes report first."}), 400
        try:
            result = enrollment_service.process(
                program=program,
                uploaded_file=uploaded,
                start_date=request.form.get("start_date", ""),
                end_date=request.form.get("end_date", ""),
                structure=service.load_structure(program),
                current_metrics=service.load_metrics(program),
                team_uploaded_file=request.files.get("team_file"),
                cls_uploaded_file=request.files.get("cls_file"),
                services_uploaded_file=request.files.get("services_file"),
            )
            service.save_structure(program, result["structure"])
            service.save_metrics(program, result["metrics"])
        except ValueError as exc:
            return jsonify({"error": str(exc)}), 400
        result["download_url"] = f"/downloads/{program.upper()}/{result['download_name']}"
        return jsonify(result)

    @app.post("/api/programs/<program>/timeliness/process")
    def process_timeliness(program):
        uploaded = request.files.get("case_notes_file")
        if not uploaded or not uploaded.filename:
            return jsonify({"error": "Drop a Case Notes report first."}), 400
        try:
            structure = service.load_structure(program)
            result = timeliness_service.process(
                program=program,
                uploaded_file=uploaded,
                start_date=request.form.get("start_date", ""),
                end_date=request.form.get("end_date", ""),
                structure=structure,
                current_metrics=service.load_metrics(program),
            )
            result.pop("raw_frame")
            result.pop("clean_frame")
            result.pop("pivots")
            service.save_metrics(program, result["metrics"])
        except ValueError as exc:
            return jsonify({"error": str(exc)}), 400
        result["download_url"] = f"/downloads/{program.upper()}/{result['download_name']}"
        return jsonify(result)

    @app.post("/api/programs/<program>/finalize")
    def finalize_report(program):
        enrollment_upload = request.files.get("enrollment_file")
        case_notes_upload = request.files.get("case_notes_file")
        if not enrollment_upload or not enrollment_upload.filename:
            return jsonify({"error": "Keep the Enrollment/Outcomes file selected before creating the complete Excel."}), 400
        if not case_notes_upload or not case_notes_upload.filename:
            return jsonify({"error": "Keep the Case Notes/Timeliness file selected before creating the complete Excel."}), 400
        try:
            structure = service.load_structure(program)
            metrics_json = request.form.get("metrics_json", "")
            current_metrics = service.load_metrics(program)
            if metrics_json:
                parsed_metrics = json.loads(metrics_json)
                if not isinstance(parsed_metrics, dict):
                    raise ValueError("Current table values must be an object keyed by staff name.")
                current_metrics = parsed_metrics

            timeliness = timeliness_service.process(
                program=program,
                uploaded_file=case_notes_upload,
                start_date=request.form.get("start_date", ""),
                end_date=request.form.get("end_date", ""),
                structure=structure,
                current_metrics=current_metrics,
            )
            combined = enrollment_service.process(
                program=program,
                uploaded_file=enrollment_upload,
                start_date=request.form.get("start_date", ""),
                end_date=request.form.get("end_date", ""),
                structure=structure,
                current_metrics=timeliness["metrics"],
                team_uploaded_file=request.files.get("team_file"),
                cls_uploaded_file=request.files.get("cls_file"),
                services_uploaded_file=request.files.get("services_file"),
                case_notes_raw=timeliness["raw_frame"],
                case_notes_clean=timeliness["clean_frame"],
                timeliness_pivots=timeliness["pivots"],
            )
            saved_metrics = service.save_metrics(program, combined["metrics"])
        except json.JSONDecodeError:
            return jsonify({"error": "The current table values could not be read."}), 400
        except ValueError as exc:
            return jsonify({"error": str(exc)}), 400

        result = {
            "metrics": saved_metrics,
            "download_name": combined["download_name"],
            "download_url": f"/downloads/{program.upper()}/{combined['download_name']}",
            "shared_metrics": combined.get("shared_metrics", {}),
            "included_sources": {
                "team": bool(request.files.get("team_file") and request.files["team_file"].filename),
                "cls": bool(request.files.get("cls_file") and request.files["cls_file"].filename),
                "services": bool(request.files.get("services_file") and request.files["services_file"].filename),
                "enrollment": True,
                "case_notes": True,
            },
        }
        return jsonify(result)

    @app.get("/downloads/<program>/<path:filename>")
    def download_output(program, filename):
        if program.upper() not in {"MHRT", "RRT"}:
            return jsonify({"error": "Unknown program."}), 404
        return send_from_directory(Path(app.config["OUTPUT_DIR"]) / program.upper(), filename, as_attachment=True)

    return app


app = create_app()


if __name__ == "__main__":
    app.run(debug=True, use_reloader=False, port=5001)
