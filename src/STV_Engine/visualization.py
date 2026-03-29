from __future__ import annotations

from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

from .models import IMPACT_KEYS, STVResults


METRIC_LABELS = {
    "carbon": "Carbon (kgCO2e)",
    "energy": "Energy (MJ)",
    "water": "Water (kgH2O)",
    "ozone": "Ozone (kgCFC11e)",
}


def export_visualizations(results: STVResults, output_dir: Path | str) -> dict[str, Path]:
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    target_vs_project_path = output_path / "target_vs_project.png"
    breakdown_path = output_path / "life_cycle_breakdown.png"
    radial_path = output_path / "target_vs_project_radial.png"

    _plot_target_vs_project(results, target_vs_project_path)
    _plot_breakdown(results, breakdown_path)
    _plot_radial_target_vs_project(results, radial_path)

    return {
        "target_vs_project": target_vs_project_path,
        "life_cycle_breakdown": breakdown_path,
        "target_vs_project_radial": radial_path,
    }


def _plot_target_vs_project(results: STVResults, output_path: Path) -> None:
    summary = results.metric_summary()
    metrics = ["carbon", "energy", "water"]
    labels = [METRIC_LABELS[key] for key in metrics]
    targets = [summary[key]["target"] or 0.0 for key in metrics]
    projects = [summary[key]["project"] or 0.0 for key in metrics]

    fig, ax = plt.subplots(figsize=(10, 5.5))
    x_positions = range(len(metrics))
    width = 0.36
    ax.bar([x - width / 2 for x in x_positions], targets, width=width, label="Target", color="#7a8fa6")
    ax.bar([x + width / 2 for x in x_positions], projects, width=width, label="Project", color="#d99058")
    ax.set_xticks(list(x_positions))
    ax.set_xticklabels(labels, rotation=10, ha="right")
    ax.set_title(f"STV Performance vs Target: {results.team}")
    ax.set_ylabel("Impact")
    ax.legend()
    ax.grid(axis="y", alpha=0.25)
    fig.tight_layout()
    fig.savefig(output_path, dpi=180)
    plt.close(fig)


def _plot_breakdown(results: STVResults, output_path: Path) -> None:
    breakdown = results.breakdown
    segments = [
        ("Materials", breakdown.embodied_materials, "#5b8e7d"),
        ("Transport", breakdown.embodied_transport, "#8aa05d"),
        ("Construction", breakdown.embodied_construction, "#d2a24c"),
        ("Electricity", breakdown.use_electricity, "#cc6b49"),
        ("Heating", breakdown.use_heating, "#8f5f99"),
        ("Water Use", breakdown.use_water, "#5d7ea8"),
    ]

    fig, ax = plt.subplots(figsize=(10, 6))
    metrics = list(IMPACT_KEYS)
    x_positions = range(len(metrics))
    bottoms = [0.0 for _ in metrics]
    for label, vector, color in segments:
        values = [vector.get(metric) for metric in metrics]
        ax.bar(x_positions, values, bottom=bottoms, label=label, color=color)
        bottoms = [bottom + value for bottom, value in zip(bottoms, values)]

    ax.set_xticks(list(x_positions))
    ax.set_xticklabels([METRIC_LABELS[key] for key in metrics], rotation=10, ha="right")
    ax.set_ylabel("Impact")
    ax.set_title(f"Life Cycle Impact Breakdown: {results.team}")
    ax.legend()
    ax.grid(axis="y", alpha=0.25)
    fig.tight_layout()
    fig.savefig(output_path, dpi=180)
    plt.close(fig)


def _plot_radial_target_vs_project(results: STVResults, output_path: Path) -> None:
    summary = results.metric_summary()
    metrics = ["carbon", "energy", "water"]
    labels = [METRIC_LABELS[key].split(" ")[0] for key in metrics]

    target_percentages = [100.0, 100.0, 100.0]
    project_percentages = [
        min((summary[key]["percent_of_target"] or 0.0) * 100.0, 200.0)
        for key in metrics
    ]

    angles = np.linspace(0, 2 * np.pi, len(metrics), endpoint=False).tolist()
    angles += angles[:1]
    target_values = target_percentages + target_percentages[:1]
    project_values = project_percentages + project_percentages[:1]

    fig, ax = plt.subplots(figsize=(7, 7), subplot_kw={"projection": "polar"})
    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(labels)
    ax.set_ylim(0, 200)
    ax.set_yticks([50, 100, 150, 200])
    ax.set_yticklabels(["50%", "100%", "150%", "200%"])
    ax.set_title(f"Target vs Project Percentage: {results.team}", pad=24)

    ax.plot(angles, target_values, color="#7a8fa6", linewidth=2, label="Target")
    ax.fill(angles, target_values, color="#7a8fa6", alpha=0.15)

    ax.plot(angles, project_values, color="#d99058", linewidth=2, label="Project")
    ax.fill(angles, project_values, color="#d99058", alpha=0.25)

    for angle, value in zip(angles[:-1], project_percentages):
        ax.text(
            angle,
            min(value + 12, 200),
            f"{value:.1f}%",
            ha="center",
            va="center",
            fontsize=9,
            color="#5a3e25",
        )

    ax.legend(loc="upper right", bbox_to_anchor=(1.15, 1.12))
    fig.tight_layout()
    fig.savefig(output_path, dpi=180)
    plt.close(fig)
