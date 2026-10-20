"""업무보고_자동취합.py와 내장 양식을 단일 EXE로 제작한다."""

import importlib.util
import sys
from pathlib import Path


def main():
    root = Path(__file__).resolve().parent

    program = root / "업무보고_자동취합.py"
    template = root / "assets" / "report_template.hwpx"

    missing = [
        path
        for path in (program, template, root / "assets" / "fonts" / "Pretendard-Regular.otf",
                     root / "assets" / "fonts" / "Pretendard-Bold.otf",
                     root / "assets" / "fonts" / "OFL.txt",
                     root / "assets" / "fonts" / "Pretendard-Medium.otf",
                     root / "assets" / "app.manifest")
        if not path.is_file()
    ]

    if missing:
        print("\nEXE 제작에 필요한 파일이 없습니다.")

        for path in missing:
            print(f"  - {path}")

        print(
            "\n업무보고_자동취합.py와 assets 폴더를 "
            "build_exe.py와 같은 폴더에 두세요."
        )
        return 1

    required_modules = {
        "PyInstaller": "pyinstaller",
        "pyhwpx": "pyhwpx",
        "lxml": "lxml",
        "win32com": "pywin32",
    }

    missing_packages = [
        package
        for module, package in required_modules.items()
        if importlib.util.find_spec(module) is None
    ]

    if missing_packages:
        print("\n필요한 제작 도구가 설치되어 있지 않습니다.")
        print("다음 명령을 실행한 뒤 다시 시도하세요.\n")
        print(
            "py -m pip install "
            + " ".join(missing_packages)
        )
        return 1

    import PyInstaller.__main__

    build_dir = root / "build"
    dist_dir = root / "dist"

    build_dir.mkdir(parents=True, exist_ok=True)
    dist_dir.mkdir(parents=True, exist_ok=True)

    exe_name = "생방기술국_회의자료_취합"
    exe_path = dist_dir / f"{exe_name}.exe"

    print("\nEXE 제작을 시작합니다.")
    print(f"프로그램: {program}")
    print(f"포함할 양식: {template}")
    print(f"결과 위치: {exe_path}\n")
    print("이전에 실행한 EXE가 있다면 먼저 닫아주세요.\n")

    options = [
        "--noconfirm",
        "--onefile",
        "--windowed",
        "--noupx",

        # 시작 시점부터 모니터 배율에 맞춰 글자를 직접 그린다.
        "--manifest",
        str(root / "assets" / "app.manifest"),

        "--name",
        exe_name,

        "--distpath",
        str(dist_dir),

        "--workpath",
        str(build_dir),

        "--specpath",
        str(build_dir),

        # 양식을 EXE 내부의 assets 폴더에 포함한다.
        "--add-data",
        f"{template}:assets",

        # 직원 PC에서 별도 설치 없이 사용하는 Pretendard 글꼴과 라이선스.
        "--add-data",
        f"{root / 'assets' / 'fonts'}:assets/fonts",

        # pyhwpx의 하위 모듈과 DLL, 자료 파일을 포함한다.
        "--collect-all",
        "pyhwpx",

        "--hidden-import",
        "win32timezone",

        "--hidden-import",
        "win32com.client.genpy",

        # 현재 취합 기능에서 사용하지 않는 선택 라이브러리.
        "--exclude-module",
        "matplotlib",

        "--exclude-module",
        "IPython",

        "--exclude-module",
        "pytest",

        "--exclude-module",
        "scipy",

        "--exclude-module",
        "sympy",

        "--exclude-module",
        "jupyter",

        str(program),
    ]

    PyInstaller.__main__.run(options)

    if not exe_path.is_file():
        print("\nEXE 파일이 생성되지 않았습니다.")
        return 1

    size_mb = exe_path.stat().st_size / (1024 * 1024)

    print("\n" + "=" * 50)
    print("EXE 제작이 완료되었습니다.")
    print(f"파일: {exe_path}")
    print(f"크기: {size_mb:.1f} MB")
    print()
    print("완성된 EXE를 직접 실행해 취합과 저장을 확인하세요.")
    print("직원에게는 확인한 EXE 파일 하나만 전달하면 됩니다.")
    print("직원 PC에는 Windows용 한/글이 필요합니다.")
    print("=" * 50)

    return 0


if __name__ == "__main__":
    sys.exit(main())