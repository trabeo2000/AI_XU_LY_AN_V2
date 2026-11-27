# -*- coding: utf-8 -*-
"""
paths_config.py - Single Source of Truth cho toàn bộ đường dẫn tương đối của hệ thống.
Khi tái cấu trúc hoặc đổi tên thư mục trong tương lai, chỉ cần cập nhật tại đây để
mọi script Python (export_dossier, export_to_trinh, export_khang_nghi, export_6step_report,
universal_visual_filler, fast_reader, sync_legal_knowledge, build_base_*, verify_system)
đồng bộ 100% tự động trên cả 3 máy tính.
"""
import os

SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
FLOW_DIR = os.path.dirname(SCRIPTS_DIR)
ROOT_DIR = os.path.dirname(FLOW_DIR)

TEMPLATES_DIR = os.path.join(FLOW_DIR, "templates")
BASE_TO_TRINH_TEMPLATE = os.path.join(TEMPLATES_DIR, "base_to_trinh_tandtc.docx")
BASE_KHANG_NGHI_TEMPLATE = os.path.join(TEMPLATES_DIR, "base_khang_nghi_tandtc.docx")
BASE_THONG_BAO_TEMPLATE = os.path.join(TEMPLATES_DIR, "base_thong_bao_tld_tandtc.docx")

TEMP_DIR = os.path.join(SCRIPTS_DIR, ".temp")
RENDERED_CACHE_DIR = os.path.join(TEMP_DIR, "rendered_cache")
REGISTRY_PATH = os.path.join(FLOW_DIR, "case_registry.json")

DOSSIER_INPUT_DIR = os.path.join(ROOT_DIR, "02_Ho_so")
OUTPUT_ROOT_DIR = os.path.join(ROOT_DIR, "03_Ket_qua_Xu_ly")
LEGAL_DOCS_DIR = os.path.join(ROOT_DIR, "04_Van_ban_phap_ly")
WORKFLOW_DOCS_DIR = os.path.join(ROOT_DIR, "05_Quy_trinh_va_Pham_vi_Chuc_danh")
AGENTS_MD_PATH = os.path.join(ROOT_DIR, ".agents", "AGENTS.md")
