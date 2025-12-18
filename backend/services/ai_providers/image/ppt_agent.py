"""
PPT Agent Image Provider
"""
import json
import os
import math
import time
import requests
import logging
from io import BytesIO
from typing import Optional, List, Dict
from PIL import Image, ImageDraw
from openai import OpenAI
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE
from .base import ImageProvider

logger = logging.getLogger(__name__)

# --- Try importing Windows COM ---
try:
    import win32com.client
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False


# ==========================================
# 🧠 Class 1: Planner Agent (策划)
# ==========================================
class PlannerAgent:
    def __init__(self, client, model="gpt-4o"):
        self.client = client
        self.model = model

    def generate_plan(self, user_input):
        logger.info(f"🧠 [Planner] Analyzing text and fusing data...")
        json_schema = """{
          "meta": {"layout_type": "string", "theme": "tech_blue"},
          "content": {
            "main_title": "string", "subtitle": "string",
            "items": [{
                "id": "string", "title": "string", "desc": "string",
                "specs": {"Key": "Value"}, "tags": ["string"]
            }]
          },
          "assets": {"images": [{"target_id": "string", "prompt": "string", "local_path": null}]}
        }"""
        system_prompt = f"""You are a high-level information architect. Fuse user input into structured data. If table/comparison data exists, it must be extracted into the items `specs` dictionary. Generate a 3D Tech Blue style image generation Prompt for each item. Output pure JSON: {json_schema}"""
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[{"role": "system", "content": system_prompt}, {"role": "user", "content": user_input}],
                temperature=0.1
            )
            content = response.choices[0].message.content
            # Clean up markdown code blocks if present
            content = content.replace("```json", "").replace("```", "").strip()
            return json.loads(content)
        except Exception as e:
            logger.error(f"❌ Planning stage failed: {e}")
            return None


# ==========================================
# 🏭 Class 2: Production Agent (生产)
# ==========================================
class ProductionAgent:
    def __init__(self, client, assets_dir, use_mock=True):
        self.client = client
        self.use_mock = use_mock
        self.assets_dir = assets_dir

    def _create_local_mock_image(self, prompt, filename):
        filepath = os.path.join(self.assets_dir, filename)
        img = Image.new('RGB', (1024, 1024), color=(10, 30, 60))
        d = ImageDraw.Draw(img)
        d.rectangle([50, 50, 974, 974], outline=(0, 255, 255), width=5)
        d.ellipse([300, 300, 724, 724], outline=(255, 255, 255), width=2)

        # Draw some text to indicate it's a mock
        try:
            # Try to use default font
            d.text((350, 500), "MOCK IMAGE", fill=(255, 255, 255))
        except:
            pass

        img.save(filepath)
        logger.info(f"🎨 [Mock] Generated local mock: {filepath}")
        return filepath

    def _generate_dalle_image(self, prompt, filename):
        logger.info(f"🎨 [DALL-E] Requesting generation: {prompt[:30]}...")
        try:
            # Note: We assume self.client is configured for DALL-E or compatible API
            # If using OpenAI proxy that supports images.generate
            response = self.client.images.generate(model="dall-e-3", prompt=prompt, size="1024x1024", n=1)
            url = response.data[0].url
            res = requests.get(url)
            if res.status_code == 200:
                filepath = os.path.join(self.assets_dir, filename)
                with open(filepath, 'wb') as f:
                    f.write(res.content)
                logger.info(f"💾 [DALL-E] Downloaded: {filepath}")
                return filepath
        except Exception as e:
            logger.error(f"❌ Image generation failed: {e}")
        return None

    def produce_assets(self, plan_data):
        logger.info(f"🏭 [Production] Starting asset production...")
        if not plan_data or 'assets' not in plan_data:
            return plan_data

        for img in plan_data.get('assets', {}).get('images', []):
            target_id = img.get('target_id', 'unknown')
            fname = f"{target_id}_{int(time.time())}.png"

            if self.use_mock:
                path = self._create_local_mock_image(img.get('prompt'), fname)
            else:
                path = self._generate_dalle_image(img.get('prompt'), fname)

            if path:
                img['local_path'] = path
        return plan_data


# ==========================================
# 🔨 Class 3: Coder Agent (渲染)
# ==========================================
class SlideRenderer:
    def __init__(self, prs, slide):
        self.slide = slide
        self.prs = prs
        self.W = prs.slide_width
        self.H = prs.slide_height
        self.C_BG = RGBColor(10, 25, 47)
        self.C_ACCENT = RGBColor(0, 255, 255)
        self.C_CARD = RGBColor(23, 42, 69)
        self.C_BORDER = RGBColor(45, 65, 95)
        self.C_TX_H = RGBColor(255, 255, 255)
        self.C_TX_B = RGBColor(170, 190, 210)

    def setup_base(self):
        bg = self.slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, self.W, self.H)
        bg.fill.solid()
        bg.fill.fore_color.rgb = self.C_BG

    def draw_header(self, title, subtitle):
        bar = self.slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.4), Inches(0.15), Inches(0.8))
        bar.fill.solid()
        bar.fill.fore_color.rgb = self.C_ACCENT
        tb = self.slide.shapes.add_textbox(Inches(0.8), Inches(0.4), self.W - Inches(1), Inches(1))
        p = tb.text_frame.paragraphs[0]
        p.text = title
        p.font.size = Pt(40)
        p.font.bold = True
        p.font.color.rgb = self.C_TX_H
        if subtitle:
            tb_s = self.slide.shapes.add_textbox(Inches(0.8), Inches(1.1), self.W - Inches(1), Inches(0.6))
            p_s = tb_s.text_frame.paragraphs[0]
            p_s.text = subtitle
            p_s.font.size = Pt(18)
            p_s.font.color.rgb = self.C_ACCENT

    def _greedy_grid(self, count, start_y):
        margin = Inches(0.5)
        gap = Inches(0.3)
        if count == 3:
            c, r = 3, 1
        elif count == 4:
            c, r = 2, 2
        elif count in [5, 6]:
            c, r = 3, 2
        else:
            c = 3
            if count == 0:
                r = 1
            else:
                r = math.ceil(count / c)

        # Avoid division by zero
        if c == 0: c = 1
        if r == 0: r = 1

        cw = (self.W - margin * 2 - gap * (c - 1)) / c
        ch = (self.H - start_y - margin - gap * (r - 1)) / r
        return [(margin + (i % c) * (cw + gap), start_y + (i // c) * (ch + gap), cw, ch) for i in range(count)]

    def render_grid(self, items, asset_map):
        if not items:
            return

        slots = self._greedy_grid(len(items), Inches(1.8))
        for i, (x, y, w, h) in enumerate(slots):
            item = items[i]
            specs = item.get('specs', {})
            has_specs = len(specs) > 0
            card = self.slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
            card.fill.solid()
            card.fill.fore_color.rgb = self.C_CARD
            card.line.color.rgb = self.C_BORDER

            isz = Inches(0.8) if has_specs else Inches(1.2)
            cy = y + Inches(0.2)
            path = asset_map.get(item.get('id'))

            tx = x + Inches(0.2)
            if path and os.path.exists(path):
                try:
                    self.slide.shapes.add_picture(path, x + Inches(0.2), cy, isz, isz)
                    tx = x + isz + Inches(0.4)
                except Exception as e:
                    logger.warning(f"Failed to add picture {path}: {e}")

            tb_t = self.slide.shapes.add_textbox(tx, cy, w - (tx - x), Inches(0.6))
            p = tb_t.text_frame.paragraphs[0]
            p.text = item.get('title', '')
            p.font.bold = True
            p.font.size = Pt(20)
            p.font.color.rgb = self.C_TX_H

            tb_d = self.slide.shapes.add_textbox(tx, cy + Inches(0.4), w - (tx - x) - Inches(0.2), Inches(0.8))
            tb_d.text_frame.word_wrap = True
            p2 = tb_d.text_frame.paragraphs[0]
            p2.text = item.get('desc', '')
            p2.font.size = Pt(14)
            p2.font.color.rgb = self.C_TX_B

            if has_specs:
                sy = cy + Inches(1.2)
                sep = self.slide.shapes.add_shape(MSO_SHAPE.LINE_INVERSE, x + Inches(0.2), sy, w - Inches(0.4), 0)
                sep.line.color.rgb = self.C_ACCENT
                sep.line.dash_style = 1

                # Prevent division by zero if specs is empty (though has_specs checks that)
                if len(specs) > 0:
                    rh = (h - (sy - y) - Inches(0.3)) / len(specs)
                    for idx, (k, v) in enumerate(specs.items()):
                        cur_y = sy + Inches(0.1) + idx * rh
                        tb_k = self.slide.shapes.add_textbox(x + Inches(0.2), cur_y, Inches(1.5), rh)
                        pk = tb_k.text_frame.paragraphs[0]
                        pk.text = f"● {k}"
                        pk.font.bold = True
                        pk.font.size = Pt(12)
                        pk.font.color.rgb = self.C_ACCENT

                        tb_v = self.slide.shapes.add_textbox(x + Inches(1.6), cur_y, w - Inches(1.8), rh)
                        tb_v.text_frame.word_wrap = True
                        tb_v.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
                        pv = tb_v.text_frame.paragraphs[0]
                        pv.text = str(v)
                        pv.font.size = Pt(12)
                        pv.font.color.rgb = RGBColor(200, 200, 200)

    def dispatch(self, final_plan, asset_map):
        self.setup_base()
        content = final_plan.get('content', {})
        self.draw_header(content.get('main_title', ''), content.get('subtitle', ''))
        self.render_grid(content.get('items', []), asset_map)


# ==========================================
# 📸 Class 4: Exporter (导出图片)
# ==========================================
class PPTExporter:
    def __init__(self):
        self.enabled = WIN32_AVAILABLE

    def export_first_slide_as_jpg(self, pptx_path):
        """Use PowerPoint background process to export the first slide as JPG"""
        abs_pptx_path = os.path.abspath(pptx_path)
        base, _ = os.path.splitext(abs_pptx_path)
        abs_jpg_path = f"{base}.jpg"

        if not self.enabled:
            logger.warning("⚠️ Skipping image export: pywin32 required (Windows only). Generating placeholder.")
            return self._generate_placeholder_image(abs_jpg_path)

        logger.info(f"📸 [Exporter] Calling PowerPoint background process...")
        logger.info(f"   - Source: {abs_pptx_path}")

        powerpoint = None
        presentation = None
        try:
            # Try dynamic dispatch first
            try:
                powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            except AttributeError:
                powerpoint = win32com.client.EnsureDispatch("PowerPoint.Application")

            try:
                powerpoint.Visible = False
            except:
                pass

            presentation = powerpoint.Presentations.Open(abs_pptx_path, ReadOnly=True, WithWindow=False)
            presentation.Slides(1).Export(abs_jpg_path, FilterName="JPG")
            logger.info(f"✅ Image exported: {abs_jpg_path}")
            return abs_jpg_path

        except Exception as e:
            logger.error(f"❌ Image export failed (COM Error): {e}")
            return self._generate_placeholder_image(abs_jpg_path, error_msg=str(e))
        finally:
            if presentation:
                try:
                    presentation.Close()
                except:
                    pass
            # We don't quit PowerPoint as it might be used by user,
            # or maybe we should if we started it?
            # Safe to just release COM object usually.
            if powerpoint:
                del powerpoint

    def _generate_placeholder_image(self, output_path, error_msg=None):
        """Generate a placeholder image when export fails (e.g. on Linux)"""
        img = Image.new('RGB', (1280, 720), color=(10, 25, 47))
        d = ImageDraw.Draw(img)
        d.rectangle([50, 50, 1230, 670], outline=(0, 255, 255), width=5)

        text = "PPT Generated Successfully"
        subtext = "Image export unavailable on this platform (Linux/No PPT)"

        # Simple centering
        d.text((400, 300), text, fill=(255, 255, 255))
        d.text((350, 350), subtext, fill=(200, 200, 200))

        if error_msg:
             d.text((100, 400), f"Error: {error_msg[:50]}...", fill=(255, 100, 100))

        img.save(output_path)
        logger.info(f"⚠️ Generated placeholder image: {output_path}")
        return output_path


class PPTAgentImageProvider(ImageProvider):
    """
    Image Provider that uses the PPT Agent workflow:
    Planner -> Production -> SlideRender -> Exporter
    """

    def __init__(self, api_key: str, api_base: str = None, model: str = "gpt-4o"):
        self.api_key = api_key
        self.api_base = api_base
        self.model = model
        # Initialize OpenAI client
        self.client = OpenAI(
            api_key=api_key,
            base_url=api_base
        )
        # Working directory for artifacts
        self.work_dir = os.path.join(os.getcwd(), "backend_assets")
        if not os.path.exists(self.work_dir):
            os.makedirs(self.work_dir)

        # Use mock images by default to save cost/time, or make it configurable?
        # The user provided code had USE_MOCK_IMAGES = True
        self.use_mock_images = True

    def generate_image(
        self,
        prompt: str,
        ref_images: Optional[List[Image.Image]] = None,
        aspect_ratio: str = "16:9",
        resolution: str = "2K"
    ) -> Optional[Image.Image]:
        """
        Generate a PPT slide image from the prompt.
        """
        try:
            timestamp = int(time.time())

            # 1. Planning
            planner = PlannerAgent(self.client, model=self.model)
            plan = planner.generate_plan(prompt)
            if not plan:
                raise ValueError("Planner failed to generate a plan.")

            # 2. Production
            # Create a subdirectory for this request to avoid collisions?
            # For simplicity, use shared assets dir
            producer = ProductionAgent(self.client, assets_dir=self.work_dir, use_mock=self.use_mock_images)
            final_plan = producer.produce_assets(plan)

            # 3. Coding & Rendering
            logger.info(f"🔨 [Coder] Building PPT...")
            prs = Presentation()
            prs.slide_width = Inches(16)
            prs.slide_height = Inches(9)
            slide = prs.slides.add_slide(prs.slide_layouts[6])

            asset_map = {
                img['target_id']: img['local_path']
                for img in final_plan.get('assets', {}).get('images', [])
                if img.get('local_path')
            }

            renderer = SlideRenderer(prs, slide)
            renderer.dispatch(final_plan, asset_map)

            # Save PPTX
            pptx_filename = f"output_{timestamp}.pptx"
            pptx_path = os.path.join(self.work_dir, pptx_filename)
            prs.save(pptx_path)
            logger.info(f"🎉 PPT Generated: {pptx_path}")

            # 4. Exporting
            exporter = PPTExporter()
            jpg_path = exporter.export_first_slide_as_jpg(pptx_path)

            if jpg_path and os.path.exists(jpg_path):
                return Image.open(jpg_path)
            else:
                raise RuntimeError("Failed to generate JPG from PPT.")

        except Exception as e:
            logger.error(f"Error in PPTAgentImageProvider: {e}", exc_info=True)
            # Optionally return None or raise
            return None
