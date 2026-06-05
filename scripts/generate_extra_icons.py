from pathlib import Path

from PIL import Image, ImageDraw


OUT_DIR = Path(__file__).resolve().parents[1] / "icons" / "google_slides_png"
OUT_DIR.mkdir(parents=True, exist_ok=True)

SIZE = 4096


def base_canvas() -> tuple[Image.Image, ImageDraw.ImageDraw]:
    img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    m = 220
    draw.rounded_rectangle(
        (m, m, SIZE - m, SIZE - m),
        radius=420,
        fill=(247, 251, 255, 255),
        outline=(186, 200, 214, 255),
        width=42,
    )
    return img, draw


def icon_compare_models() -> None:
    img, d = base_canvas()
    cx, cy = SIZE // 2, SIZE // 2 + 120
    pts = [(cx - 900, cy - 380), (cx, cy - 620), (cx + 900, cy - 380), (cx - 580, cy + 520), (cx + 580, cy + 520)]
    cols = [(42, 108, 246, 255), (25, 167, 144, 255), (244, 128, 36, 255), (123, 97, 255, 255), (58, 179, 83, 255)]

    for x, y in pts:
        d.line((cx, cy, x, y), fill=(140, 160, 180, 255), width=28)
    for (x, y), c in zip(pts, cols):
        d.ellipse((x - 250, y - 250, x + 250, y + 250), fill=(225, 236, 248, 255), outline=c, width=56)
        d.line((x - 90, y, x + 90, y), fill=c, width=30)
        d.line((x, y - 90, x, y + 90), fill=c, width=30)

    d.ellipse((cx - 170, cy - 170, cx + 170, cy + 170), fill=(39, 48, 61, 255))
    d.ellipse((cx - 80, cy - 80, cx + 80, cy + 80), fill=(247, 251, 255, 255))
    img.save(OUT_DIR / "09_compare_five_model_families_4096.png")


def icon_runtime_energy_proxy() -> None:
    img, d = base_canvas()
    # clock
    d.ellipse((540, 1190, 2060, 2710), fill=(225, 236, 248, 255), outline=(42, 108, 246, 255), width=62)
    d.line((1300, 1950, 1300, 1590), fill=(42, 108, 246, 255), width=54)
    d.line((1300, 1950, 1600, 2070), fill=(42, 108, 246, 255), width=54)

    # x sign
    d.line((2100, 1730, 2360, 1990), fill=(39, 48, 61, 255), width=64)
    d.line((2360, 1730, 2100, 1990), fill=(39, 48, 61, 255), width=64)

    # energy + residual bars
    d.polygon([(2720, 1300), (2550, 1720), (2830, 1720), (2640, 2400), (3300, 1650), (2970, 1650)], fill=(244, 128, 36, 255), outline=(215, 100, 24, 255))
    d.rounded_rectangle((2520, 2520, 3380, 3120), radius=120, fill=(235, 245, 236, 255), outline=(58, 179, 83, 255), width=40)
    for x, h in zip([2620, 2810, 3000, 3190], [280, 430, 220, 510]):
        d.rounded_rectangle((x, 3040 - h, x + 120, 3040), radius=25, fill=(25, 167, 144, 255))

    img.save(OUT_DIR / "10_runtime_energy_residual_proxy_4096.png")


def icon_pareto_candidates() -> None:
    img, d = base_canvas()
    d.rounded_rectangle((700, 820, 3390, 3200), radius=80, fill=(250, 252, 255, 255), outline=(180, 195, 210, 255), width=32)
    d.line((940, 2940, 3140, 2940), fill=(39, 48, 61, 255), width=34)
    d.line((940, 2940, 940, 1060), fill=(39, 48, 61, 255), width=34)

    points = [(1200, 2550), (1450, 2250), (1780, 1900), (2170, 1740), (2480, 1500), (2830, 1360)]
    for i, (x, y) in enumerate(points):
        c = (58, 179, 83, 255) if i >= 3 else (222, 76, 94, 255)
        d.ellipse((x - 70, y - 70, x + 70, y + 70), fill=c, outline=(39, 48, 61, 255), width=18)

    d.line(points[2:6], fill=(123, 97, 255, 255), width=28)
    x, y = points[-1]
    d.ellipse((x - 150, y - 150, x + 150, y + 150), outline=(244, 128, 36, 255), width=36)
    img.save(OUT_DIR / "11_pareto_low_carbon_candidates_4096.png")


def icon_schedule_training() -> None:
    img, d = base_canvas()
    d.rounded_rectangle((620, 2400, 3470, 3000), radius=100, fill=(242, 247, 252, 255), outline=(176, 192, 209, 255), width=30)
    for i in range(13):
        x = 700 + i * 220
        d.line((x, 2450, x, 2950), fill=(165, 180, 198, 255), width=8)
    d.rounded_rectangle((1820, 2440, 2480, 2960), radius=70, fill=(209, 239, 214, 255), outline=(58, 179, 83, 255), width=18)

    d.rounded_rectangle((1460, 1080, 2620, 1900), radius=120, fill=(225, 236, 248, 255), outline=(42, 108, 246, 255), width=44)
    for x in [1520, 2560]:
        for y in [1220, 1400, 1580, 1760]:
            d.rounded_rectangle((x - 60, y - 30, x + 60, y + 30), radius=10, fill=(42, 108, 246, 255))
    for x in [1760, 1980, 2200, 2420]:
        for y in [1140, 1840]:
            d.rounded_rectangle((x - 30, y - 60, x + 30, y + 60), radius=10, fill=(42, 108, 246, 255))

    d.line((1670, 1690, 1880, 1510, 2060, 1600, 2270, 1360, 2420, 1460), fill=(123, 97, 255, 255), width=24)
    d.line((2048, 1930, 2048, 2360), fill=(39, 48, 61, 255), width=52)
    d.polygon([(1918, 2250), (2178, 2250), (2048, 2420)], fill=(39, 48, 61, 255))
    img.save(OUT_DIR / "12_schedule_training_low_residual_hours_4096.png")


def main() -> None:
    icon_compare_models()
    icon_runtime_energy_proxy()
    icon_pareto_candidates()
    icon_schedule_training()
    print("Created 4 icons in:", OUT_DIR)


if __name__ == "__main__":
    main()
