"""
OCREngine - Tesseract OCR 모듈
검색 결과 텍스트 인식 및 이름 매칭
"""

try:
    import pytesseract
    from PIL import Image, ImageFilter, ImageEnhance
except ImportError:
    pytesseract = None
    Image = None


class OCREngine:
    """Tesseract 기반 OCR 엔진"""

    def __init__(self, language: str = "kor+eng", confidence_threshold: int = 70):
        self.language = language
        self.confidence_threshold = confidence_threshold
        self.available = False
        self.tessdata_dir = None
        self._check_tesseract()

    def _check_tesseract(self):
        """Tesseract 설치 확인 + 프로젝트 내 tessdata 우선 사용"""
        if pytesseract is None:
            return

        import platform
        import os

        # Windows: 설치 경로 설정
        if platform.system() == "Windows":
            tesseract_path = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
            if os.path.exists(tesseract_path):
                pytesseract.pytesseract.tesseract_cmd = tesseract_path

        # 프로젝트 내 tessdata 확인 (exe와 스크립트 모두 지원)
        import sys
        if getattr(sys, 'frozen', False):
            # exe: exe 옆 config/tessdata 또는 _internal/config/tessdata
            exe_dir = os.path.dirname(sys.executable)
            candidates = [
                os.path.join(exe_dir, "config", "tessdata"),
                os.path.join(getattr(sys, '_MEIPASS', ''), "config", "tessdata"),
                os.path.join(exe_dir, "_internal", "config", "tessdata"),
            ]
        else:
            candidates = [
                os.path.join(os.path.dirname(os.path.dirname(__file__)), "config", "tessdata")
            ]

        for path in candidates:
            if os.path.exists(os.path.join(path, "kor.traineddata")):
                self.tessdata_dir = path
                break

        try:
            pytesseract.get_tesseract_version()
            self.available = True
        except Exception:
            self.available = False

    def _get_config(self, psm: int) -> str:
        """Tesseract config 문자열 생성 (tessdata 경로 포함)"""
        config = f"--psm {psm}"
        if self.tessdata_dir:
            # 경로를 슬래시로 변환 (Windows 백슬래시 문제 방지)
            td = self.tessdata_dir.replace("\\", "/")
            config += f" --tessdata-dir {td}"
        return config

    def preprocess_image(self, image: "Image.Image") -> "Image.Image":
        """OCR 정확도 향상을 위한 이미지 전처리 (다크모드 대응)"""
        # 그레이스케일 변환
        gray = image.convert("L")

        # 다크모드 감지: 평균 밝기가 낮으면 반전
        from PIL import ImageOps, ImageStat
        avg_brightness = ImageStat.Stat(gray).mean[0]
        if avg_brightness < 128:
            gray = ImageOps.invert(gray)

        # 대비 향상
        enhancer = ImageEnhance.Contrast(gray)
        enhanced = enhancer.enhance(2.0)

        # 샤프닝
        sharpened = enhanced.filter(ImageFilter.SHARPEN)

        # 3배 확대 (OCR 정확도 향상)
        width, height = sharpened.size
        resized = sharpened.resize((width * 3, height * 3), Image.LANCZOS)

        # 이진화 (흑백 분리 - 텍스트 선명도 극대화)
        threshold = 180
        resized = resized.point(lambda p: 255 if p > threshold else 0)

        return resized

    def extract_text(self, image: "Image.Image", preprocess: bool = True) -> str:
        """이미지에서 텍스트 추출 - 여러 PSM 모드 시도"""
        if not self.available:
            return ""

        if preprocess:
            image = self.preprocess_image(image)

        # 여러 PSM 모드로 시도하여 가장 많은 텍스트를 추출
        best_text = ""
        for psm in [11, 3, 6]:  # sparse → auto → uniform block
            try:
                text = pytesseract.image_to_string(
                    image,
                    lang=self.language,
                    config=self._get_config(psm)
                ).strip()
                if len(text) > len(best_text):
                    best_text = text
            except Exception:
                continue
        return best_text

    def extract_text_with_data(self, image: "Image.Image", preprocess: bool = True) -> list:
        """텍스트와 위치/신뢰도 정보 함께 추출"""
        if not self.available:
            return []

        if preprocess:
            image = self.preprocess_image(image)

        all_results = []
        # 여러 PSM 모드로 시도
        for psm in [11, 3, 6]:
            try:
                data = pytesseract.image_to_data(
                    image,
                    lang=self.language,
                    config=self._get_config(psm),
                    output_type=pytesseract.Output.DICT
                )

                n = len(data["text"])
                for i in range(n):
                    text = data["text"][i].strip()
                    conf = int(data["conf"][i])
                    if text and conf >= self.confidence_threshold:
                        # 중복 방지
                        if not any(r["text"] == text for r in all_results):
                            all_results.append({
                                "text": text,
                                "confidence": conf,
                                "x": data["left"][i],
                                "y": data["top"][i],
                                "width": data["width"][i],
                                "height": data["height"][i]
                            })
            except Exception:
                continue
        return all_results

    def verify_name_in_results(self, image: "Image.Image", target_name: str) -> dict:
        """
        검색 결과 이미지에서 대상 이름이 있는지 확인

        Returns:
            {
                "found": bool,
                "matched_text": str,
                "confidence": int,
                "position": {"x": int, "y": int}
            }
        """
        if not self.available:
            return {
                "found": False,
                "matched_text": None,
                "confidence": 0,
                "position": None,
                "error": "Tesseract OCR이 설치되지 않았습니다."
            }

        # 1단계: 신뢰도 높은 결과에서 매칭
        results = self.extract_text_with_data(image)

        for r in results:
            if target_name in r["text"] or r["text"] in target_name:
                return {
                    "found": True,
                    "matched_text": r["text"],
                    "confidence": r["confidence"],
                    "position": {"x": r["x"], "y": r["y"]}
                }

        # 부분 매칭 (2글자 이상 일치)
        for r in results:
            common_chars = set(target_name) & set(r["text"])
            if len(common_chars) >= 2 and len(r["text"]) >= 2:
                return {
                    "found": True,
                    "matched_text": r["text"],
                    "confidence": r["confidence"],
                    "position": {"x": r["x"], "y": r["y"]},
                    "partial_match": True
                }

        # 2단계: 전체 텍스트에서 검색 (신뢰도 무시)
        full_text = self.extract_text(image)
        if target_name in full_text:
            return {
                "found": True,
                "matched_text": target_name,
                "confidence": 50,
                "position": None
            }

        # 3단계: 신뢰도 기준 낮추고 재시도 (confidence 0 이상 모두)
        preprocessed = self.preprocess_image(image)
        for psm in [11, 3, 6]:
            try:
                data = pytesseract.image_to_data(
                    preprocessed,
                    lang=self.language,
                    config=self._get_config(psm),
                    output_type=pytesseract.Output.DICT
                )
                all_text = " ".join(
                    t.strip() for t in data["text"] if t.strip()
                )
                if target_name in all_text:
                    return {
                        "found": True,
                        "matched_text": target_name,
                        "confidence": 30,
                        "position": None,
                        "low_confidence": True
                    }
                # 각 글자가 모두 포함되어 있는지
                name_chars = set(target_name)
                text_chars = set(all_text.replace(" ", ""))
                if name_chars and name_chars.issubset(text_chars):
                    return {
                        "found": True,
                        "matched_text": all_text[:20],
                        "confidence": 20,
                        "position": None,
                        "char_match": True
                    }
            except Exception:
                continue

        return {
            "found": False,
            "matched_text": None,
            "confidence": 0,
            "position": None,
            "extracted_text": full_text
        }
