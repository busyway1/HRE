#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
VBA UTF-8 → CP949 변환 스크립트 (Windows 전용)
사용법: python convert_to_cp949_windows.py
"""

import os
import glob
import sys

def convert_to_cp949(filepath):
    """UTF-8 VBA 파일을 CP949로 변환"""
    try:
        # UTF-8로 읽기
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # CP949에서 지원하지 않는 문자 치환
        replacements = {
            '©': '(c)',
            '™': '(tm)',
            '®': '(r)',
            '•': '*',
            '—': '--',
            '–': '-',
            '"': '"',
            '"': '"',
            ''': "'",
            ''': "'"
        }
        
        for old, new in replacements.items():
            content = content.replace(old, new)
        
        # CP949로 인코딩 테스트
        try:
            content.encode('cp949')
        except UnicodeEncodeError as e:
            print(f"[오류] {os.path.basename(filepath)}: {e}")
            return False
        
        # CP949로 저장
        with open(filepath, 'w', encoding='cp949') as f:
            f.write(content)
        
        print(f"[완료] {os.path.basename(filepath)}")
        return True
        
    except Exception as e:
        print(f"[실패] {filepath}: {e}")
        return False

def main():
    """모든 UTF-8 VBA 파일을 CP949로 변환"""
    
    # UTF-8 파일 목록 (한글 포함 파일만)
    utf8_files = [
        "현재_통합_문서_code.bas",
        "CoAMaster_code.bas",
        "mod_01_FilterSearch.bas",
        "mod_02_FilterSearch_Master.bas",
        "mod_03_PTB_CoA_Input.bas",
        "mod_04_IntializeProgress.bas",
        "mod_05_PTB_Highlight.bas",
        "mod_06_VerifySum.bas",
        "mod_09_CheckMaster.bas",
        "mod_10_Public.bas",
        "mod_11_Sync.bas",
        "mod_16_Export.bas",
        "mod_17_ExchangeRate.bas",
        "mod_Log.bas",
        "mod_OpenPage.bas",
        "mod_QueryProtection.bas",
        "mod_Refresh.bas",
        "mod_Ribbon.bas",
        "mod_z_Module_GetCursor.bas",
        "Module1.bas",
        "Setup_CoAMaster.bas"
    ]
    
    print("=" * 60)
    print("VBA UTF-8 → CP949 변환 시작")
    print("=" * 60)
    print()
    
    success_count = 0
    fail_count = 0
    
    for filename in utf8_files:
        if os.path.exists(filename):
            if convert_to_cp949(filename):
                success_count += 1
            else:
                fail_count += 1
        else:
            print(f"[건너뜀] {filename}: 파일 없음")
    
    print()
    print("=" * 60)
    print(f"변환 완료: {success_count}개 성공, {fail_count}개 실패")
    print("=" * 60)
    
    if fail_count > 0:
        sys.exit(1)
    
    print()
    print("다음 단계:")
    print("1. Excel 파일 열기")
    print("2. Alt+F11 → VBA Editor")
    print("3. 파일 → 파일 가져오기")
    print("4. 변환된 .bas 파일들 선택")
    print("5. 기존 모듈 교체 확인")

if __name__ == "__main__":
    main()
