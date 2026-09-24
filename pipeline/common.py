# -*- coding: utf-8 -*-
"""원본 데이터 경로 해석 유틸.

macOS 파일시스템은 한글 파일명을 NFD로 저장하므로, 파이썬 문자열(NFC)과
직접 비교하면 매칭에 실패합니다. 모든 이름 비교는 NFC 정규화 후 수행합니다.
"""
import os, unicodedata

# 이 스크립트는 <repo>/pipeline/ 에 위치
ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))

def N(s):
    return unicodedata.normalize("NFC", s)

def ls(p):
    return sorted(os.listdir(p))

def find_file(parent, needle):
    """parent 안에서 이름에 needle(NFC)이 포함된 첫 항목의 전체 경로."""
    for d in ls(parent):
        if needle in N(d):
            return os.path.join(parent, d)
    raise FileNotFoundError(f"{needle} in {parent}")

find_dir = find_file  # 이름만 다른 동일 동작

# 주요 경로
RAW_W1  = os.path.join(ROOT, "data", "raw", "wave1_2021", "spss")
RAW_W2  = os.path.join(ROOT, "data", "raw", "wave2_2025", "spss")
MAPPING = os.path.join(ROOT, "data", "mapping")
DELIVERY = os.path.join(ROOT, "delivery", "업체전달_데이터셋_v2")
BASE = ROOT  # 하위 호환
