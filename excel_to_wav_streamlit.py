import streamlit as st
import asyncio
import edge_tts
from edge_tts import list_voices
import openpyxl
import os
import json
from pydub import AudioSegment
import io

class ExcelToWavConverterStreamlit:
    def __init__(self):
        st.title("Excel to WAV 변환기")

        # 설정 파일 경로 (Streamlit에서는 세션 상태를 사용하는 것이 더 적절할 수 있습니다)
        self.config_file = "converter_config.json" # 더 이상 파일에 저장하지 않음

        # 변수 초기화 - 기본값 설정 (세션 상태 사용)
        if 'excel_path' not in st.session_state:
            st.session_state.excel_path = ''
        if 'output_dir' not in st.session_state:
            st.session_state.output_dir = ''
        if 'start_row' not in st.session_state:
            st.session_state.start_row = 1
        if 'end_row' not in st.session_state:
            st.session_state.end_row = 10
        if 'eng_voice' not in st.session_state:
            st.session_state.eng_voice = "en-US-SteffanNeural"
        if 'kor_voice' not in st.session_state:
            st.session_state.kor_voice = "ko-KR-SunHiNeural"
        if 'chn_voice' not in st.session_state:
            st.session_state.chn_voice = "zh-CN-YunxiNeural"
        if 'eng_speed' not in st.session_state:
            st.session_state.eng_speed = 1.0
        if 'kor_speed' not in st.session_state:
            st.session_state.kor_speed = 1.0
        if 'chn_speed' not in st.session_state:
            st.session_state.chn_speed = 1.0
        if 'selected_language' not in st.session_state:
            st.session_state.selected_language = 'en'
        if 'convert_to_mp3' not in st.session_state:
            st.session_state.convert_to_mp3 = False
        if 'voices_loaded' not in st.session_state:
            st.session_state.voices_loaded = False

        self.create_widgets()
        self.load_voices() # Streamlit 앱 시작 시 음성 목록 로드

    def load_voices(self):
        """음성 목록 로드 (Streamlit 앱 시작 시 한 번만)"""
        if not st.session_state.voices_loaded:
            async def load_voices_async():
                try:
                    voices = await list_voices()
                    st.session_state.eng_voices = [v["ShortName"] for v in voices if v["Locale"].startswith("en-")]
                    st.session_state.kor_voices = [v["ShortName"] for v in voices if v["Locale"].startswith("ko-")]
                    st.session_state.chn_voices = [v["ShortName"] for v in voices if v["Locale"].startswith("zh-")]

                    if st.session_state.eng_voice not in st.session_state.eng_voices and st.session_state.eng_voices:
                        st.session_state.eng_voice = st.session_state.eng_voices[0]
                    if st.session_state.kor_voice not in st.session_state.kor_voices and st.session_state.kor_voices:
                        st.session_state.kor_voice = st.session_state.kor_voices[0]
                    if st.session_state.chn_voice not in st.session_state.chn_voices and st.session_state.chn_voices:
                        st.session_state.chn_voice = st.session_state.chn_voices[0]

                    st.session_state.voices_loaded = True # 음성 목록 로드 완료 상태 업데이트
                except Exception as e:
                    st.error(f"음성 목록을 불러오는 중 오류 발생: {str(e)}")
            asyncio.run(load_voices_async())

    def create_widgets(self):
        """Streamlit 위젯 생성"""
        # 파일 선택 섹션
        with st.expander("파일 선택", expanded=True):
            uploaded_file = st.file_uploader("엑셀 파일 선택", type=["xlsx"])
            if uploaded_file:
                st.session_state.excel_path = uploaded_file

        # 언어 및 범위 설정 섹션을 한 줄로 표시
        with st.expander("언어 및 범위 설정", expanded=True):
            col1, col2, col3, col4 = st.columns([1, 1, 1, 1])
            with col1:
                language = st.radio("언어 선택", ["en", "ko", "zh"], index=["en", "ko", "zh"].index(st.session_state.selected_language), key="lang_radio")
                st.session_state.selected_language = language
            with col2:
                start_row = st.number_input("시작 행", min_value=1, value=st.session_state.start_row, key="start_row_input")
                st.session_state.start_row = start_row
            with col3:
                end_row = st.number_input("끝 행", min_value=start_row, value=st.session_state.end_row, key="end_row_input")
                st.session_state.end_row = end_row

        # 음성 및 속도 설정 섹션을 한 줄로 표시
        with st.expander("음성 및 속도 설정", expanded=True):
            if st.session_state.voices_loaded: # 음성 목록이 로드된 경우에만 표시
                col5, col6 = st.columns(2)
                with col5:
                    if st.session_state.selected_language == 'en':
                        voice_options = st.session_state.eng_voices
                        default_voice_index = voice_options.index(st.session_state.eng_voice) if st.session_state.eng_voice in voice_options else 0
                        selected_voice = st.selectbox("영어 음성 선택", voice_options, index=default_voice_index, key="eng_voice_select")
                        st.session_state.eng_voice = selected_voice
                    elif st.session_state.selected_language == 'ko':
                        voice_options = st.session_state.kor_voices
                        default_voice_index = voice_options.index(st.session_state.kor_voice) if st.session_state.kor_voice in voice_options else 0
                        selected_voice = st.selectbox("한국어 음성 선택", voice_options, index=default_voice_index, key="kor_voice_select")
                        st.session_state.kor_voice = selected_voice
                    elif st.session_state.selected_language == 'zh':
                        voice_options = st.session_state.chn_voices
                        default_voice_index = voice_options.index(st.session_state.chn_voice) if st.session_state.chn_voice in voice_options else 0
                        selected_voice = st.selectbox("중국어 음성 선택", voice_options, index=default_voice_index, key="chn_voice_select")
                        st.session_state.chn_voice = selected_voice

                with col6:
                    if st.session_state.selected_language == 'en':
                        speed = st.slider("영어 속도 조절", 0.5, 2.0, st.session_state.eng_speed, step=0.1, key="eng_speed_slider")
                        st.session_state.eng_speed = speed
                        if st.button("영어 음성 샘플 재생", key="play_eng_sample"):
                            self.play_sample("en")

                    elif st.session_state.selected_language == 'ko':
                        speed = st.slider("한국어 속도 조절", 0.5, 2.0, st.session_state.kor_speed, step=0.1, key="kor_speed_slider")
                        st.session_state.kor_speed = speed
                        if st.button("한국어 음성 샘플 재생", key="play_kor_sample"):
                            self.play_sample("ko")

                    elif st.session_state.selected_language == 'zh':
                        speed = st.slider("중국어 속도 조절", 0.5, 2.0, st.session_state.chn_speed, step=0.1, key="chn_speed_slider")
                        st.session_state.chn_speed = speed
                        if st.button("중국어 음성 샘플 재생", key="play_chn_sample"):
                            self.play_sample("zh")
            else:
                st.text("음성 목록 로딩 중...")

        # 파일 생성 섹션
        with st.expander("파일 생성 옵션", expanded=True):
            col3, col4 = st.columns([3, 1])
            with col3:
                output_format = col3.radio("파일 형식 선택", ["WAV", "MP3"], key="format_radio", index=0) # 기본 WAV 선택
            with col4:
                mp3_conversion = st.checkbox("MP3 파일로 변환", value=False, key="mp3_checkbox", disabled= output_format == "MP3") # MP3 선택시 체크박스 비활성화
                if output_format == "MP3":
                    st.session_state.convert_to_mp3 = True # MP3 라디오 버튼 선택시 MP3 변환 활성화
                else:
                    st.session_state.convert_to_mp3 = mp3_conversion # 체크박스 값에 따라 MP3 변환 설정

            if st.button(f"{output_format} 파일 생성", key="convert_button"):
                if uploaded_file is None:
                    st.error("엑셀 파일을 먼저 선택해주세요.")
                else:
                    if output_format == "WAV":
                        self.convert_files(wav=True, mp3=False) # WAV 파일만 생성
                    elif output_format == "MP3":
                        self.convert_files(wav=False, mp3=True) # MP3 파일만 생성 (내부적으로 WAV 생성 후 MP3 변환)

    async def convert_single_word(self, text, voice, speed, output_path):
        try:
            voice_lang = voice.split("-")[0].lower()
            current_lang = st.session_state.selected_language

            if voice_lang != current_lang:
                raise ValueError(f"선택한 언어({current_lang})와 음성 언어({voice_lang})가 일치하지 않습니다.")

            if current_lang == "ko" and not self.is_korean(text):
                raise ValueError("한국어 텍스트가 아닙니다.")
            elif current_lang == "zh" and not self.is_chinese(text):
                raise ValueError("중국어 텍스트가 아닙니다.")

            if not text or not isinstance(text, str):
                raise ValueError("텍스트가 유효하지 않습니다.")
            if not voice or not isinstance(voice, str):
                raise ValueError("음성 설정이 유효하지 않습니다.")
            if not (0.5 <= speed <= 2.0):
                raise ValueError("속도는 0.5에서 2.0 사이여야 합니다.")

            communicate = edge_tts.Communicate(
                text=text,
                voice=voice,
                rate=f"+{int((speed-1)*100)}%"
            )

            temp_path = output_path + ".tmp"
            await communicate.save(temp_path)

            if os.path.exists(temp_path) and os.path.getsize(temp_path) > 0:
                os.rename(temp_path, output_path)
                return True
            else:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
                raise Exception("생성된 오디오 파일이 비어 있거나 유효하지 않습니다.")

        except Exception as e:
            print(f"오디오 생성 실패: {str(e)}")
            print(f"입력값 - 텍스트: {text}, 음성: {voice}, 속도: {speed}")
            if 'temp_path' in locals() and os.path.exists(temp_path):
                os.remove(temp_path)
            raise Exception(f"오디오 생성 중 오류 발생: {str(e)}")

    def ensure_directories(self):
        base_dir = "streamlit_output" # Streamlit 앱의 기본 출력 폴더
        os.makedirs(base_dir, exist_ok=True)
        return base_dir

    def convert_to_mp3_file(self, wav_path):
        """WAV 파일을 MP3로 변환"""
        try:
            mp3_path = wav_path.replace('.wav', '.mp3')
            audio = AudioSegment.from_wav(wav_path)
            audio.export(mp3_path, format='mp3', bitrate='192k')
            return True
        except Exception as e:
            st.error(f"MP3 변환 중 오류 발생: {e}")
            return False

    async def convert_to_wav_async(self):
        try:
            excel_file = st.session_state.excel_path
            if not excel_file:
                raise ValueError("엑셀 파일을 선택해주세요.")

            output_dir = self.ensure_directories() # 출력 디렉토리 설정

            wb = openpyxl.load_workbook(excel_file)
            sheet = wb.active

            start = int(st.session_state.start_row)
            end = int(st.session_state.end_row)
            lang = st.session_state.selected_language

            progress_bar = st.progress(0) # Streamlit progress bar
            total_rows = end - start + 1
            status_text = st.empty() # 상태 텍스트 표시 위한 empty placeholder

            output_files = [] # 생성된 파일 목록 저장

            for i, row_num in enumerate(range(start, end + 1)):
                if lang == "en":
                    text = str(sheet[f'A{row_num}'].value) if sheet[f'A{row_num}'].value else None
                    voice = st.session_state.eng_voice
                    speed = st.session_state.eng_speed
                    file_prefix = "en"
                elif lang == "ko":
                    text = str(sheet[f'B{row_num}'].value) if sheet[f'B{row_num}'].value else None
                    voice = st.session_state.kor_voice
                    speed = st.session_state.kor_speed
                    file_prefix = "ko"
                elif lang == "zh":
                    text = str(sheet[f'C{row_num}'].value) if sheet[f'C{row_num}'].value else None
                    voice = st.session_state.chn_voice
                    speed = st.session_state.chn_speed
                    file_prefix = "zh"
                else:
                    text = None

                if text:
                    output_filename = f"{file_prefix}{row_num}.wav"
                    output_path = os.path.join(output_dir, output_filename)
                    status_text.text(f"처리 중: {row_num}행, 파일명: {output_filename}") # 상태 텍스트 업데이트
                    try:
                        await self.convert_single_word(text, voice, speed, output_path)
                        output_files.append(output_path) # 파일 목록에 추가
                        if st.session_state.convert_to_mp3:
                            self.convert_to_mp3_file(output_path)
                    except Exception as e:
                        st.error(f"행 {row_num} 처리 중 오류 발생: {e}")

                progress_percent = int((i + 1) / total_rows * 100)
                progress_bar.progress(progress_percent)

            status_text.text("변환 완료!") # 최종 상태 텍스트 업데이트

            # 다운로드 버튼 생성 (zip 파일로 묶어서 다운로드)
            if output_files:
                self.create_download_zip(output_files, output_dir)

        except Exception as e:
            st.error(str(e))
            status_text.text("오류 발생") # 오류 발생 상태 텍스트 업데이트

    def create_download_zip(self, file_paths, output_dir):
        """생성된 파일들을 zip으로 묶어 다운로드 버튼 생성"""
        import zipfile
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
            for file_path in file_paths:
                arcname = os.path.basename(file_path) # zip 파일 내에 저장될 이름
                zf.write(file_path, arcname=arcname)
        zip_buffer.seek(0)

        st.download_button(
            label="다운로드 ZIP 파일",
            data=zip_buffer,
            file_name="audio_files.zip",
            mime="application/zip",
            key="download_zip_button"
        )

    def convert_to_wav(self):
        asyncio.run(self.convert_to_wav_async())

    async def play_sample_async(self, lang):
        try:
            sample_texts = {
                "en": "This is a sample voice. How does it sound?",
                "ko": "샘플 음성입니다. 어떻게 들리나요?",
                "zh": "这是示例语音。听起来怎么样？"
            }

            voice_vars = {
                "en": (st.session_state.eng_voice, st.session_state.eng_speed),
                "ko": (st.session_state.kor_voice, st.session_state.kor_speed),
                "zh": (st.session_state.chn_voice, st.session_state.chn_speed)
            }

            voice, speed = voice_vars[lang]
            text = sample_texts[lang]

            temp_dir = "streamlit_temp" # Streamlit 임시 폴더
            os.makedirs(temp_dir, exist_ok=True)
            temp_file = os.path.join(temp_dir, f"sample_{lang}.wav")

            success = await self.convert_single_word(text, voice, speed, temp_file)
            if not success:
                raise Exception("샘플 오디오 생성 실패")

            if os.path.exists(temp_file):
                with open(temp_file, 'rb') as f:
                    audio_bytes = f.read()
                st.audio(audio_bytes, format='audio/wav')
                os.remove(temp_file)

        except Exception as e:
            st.error(f"샘플 재생 중 오류 발생: {str(e)}")

    def play_sample(self, lang):
        asyncio.run(self.play_sample_async(lang))

    def convert_files(self, wav=True, mp3=True):
        """WAV 또는 MP3 파일 생성"""
        if wav or mp3:
            self.convert_to_wav()

    def is_korean(self, text):
        return any(0xAC00 <= ord(c) <= 0xD7A3 for c in text)

    def is_chinese(self, text):
        return any(0x4E00 <= ord(c) <= 0x9FFF for c in text)

if __name__ == "__main__":
    app = ExcelToWavConverterStreamlit() 