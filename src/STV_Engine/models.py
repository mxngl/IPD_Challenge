from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any


IMPACT_KEYS = ("carbon", "energy", "water", "ozone")


@dataclass(slots=True)
class ImpactVector:
    carbon: float = 0.0
    energy: float = 0.0
    water: float = 0.0
    ozone: float = 0.0

    def __add__(self, other: "ImpactVector") -> "ImpactVector":
        return ImpactVector(
            carbon=self.carbon + other.carbon,
            energy=self.energy + other.energy,
            water=self.water + other.water,
            ozone=self.ozone + other.ozone,
        )

    def scale(self, factor: float) -> "ImpactVector":
        return ImpactVector(
            carbon=self.carbon * factor,
            energy=self.energy * factor,
            water=self.water * factor,
            ozone=self.ozone * factor,
        )

    def get(self, key: str) -> float:
        return getattr(self, key)

    def to_dict(self) -> dict[str, float]:
        return {key: self.get(key) for key in IMPACT_KEYS}

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> "ImpactVector":
        return cls(
            carbon=float(payload.get("carbon", 0.0)),
            energy=float(payload.get("energy", 0.0)),
            water=float(payload.get("water", 0.0)),
            ozone=float(payload.get("ozone", 0.0)),
        )


@dataclass(slots=True)
class ImpactBreakdown:
    embodied_materials: ImpactVector = field(default_factory=ImpactVector)
    embodied_transport: ImpactVector = field(default_factory=ImpactVector)
    embodied_construction: ImpactVector = field(default_factory=ImpactVector)
    use_electricity: ImpactVector = field(default_factory=ImpactVector)
    use_heating: ImpactVector = field(default_factory=ImpactVector)
    use_water: ImpactVector = field(default_factory=ImpactVector)

    @property
    def embodied(self) -> ImpactVector:
        return (
            self.embodied_materials
            + self.embodied_transport
            + self.embodied_construction
        )

    @property
    def use_phase(self) -> ImpactVector:
        return self.use_electricity + self.use_heating + self.use_water

    @property
    def life_cycle(self) -> ImpactVector:
        return self.embodied + self.use_phase

    def __add__(self, other: "ImpactBreakdown") -> "ImpactBreakdown":
        return ImpactBreakdown(
            embodied_materials=self.embodied_materials + other.embodied_materials,
            embodied_transport=self.embodied_transport + other.embodied_transport,
            embodied_construction=self.embodied_construction + other.embodied_construction,
            use_electricity=self.use_electricity + other.use_electricity,
            use_heating=self.use_heating + other.use_heating,
            use_water=self.use_water + other.use_water,
        )

    def to_dict(self) -> dict[str, dict[str, float]]:
        return {
            "embodied_materials": self.embodied_materials.to_dict(),
            "embodied_transport": self.embodied_transport.to_dict(),
            "embodied_construction": self.embodied_construction.to_dict(),
            "use_electricity": self.use_electricity.to_dict(),
            "use_heating": self.use_heating.to_dict(),
            "use_water": self.use_water.to_dict(),
            "embodied": self.embodied.to_dict(),
            "use_phase": self.use_phase.to_dict(),
            "life_cycle": self.life_cycle.to_dict(),
        }

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> "ImpactBreakdown":
        return cls(
            embodied_materials=ImpactVector.from_dict(payload.get("embodied_materials", {})),
            embodied_transport=ImpactVector.from_dict(payload.get("embodied_transport", {})),
            embodied_construction=ImpactVector.from_dict(payload.get("embodied_construction", {})),
            use_electricity=ImpactVector.from_dict(payload.get("use_electricity", {})),
            use_heating=ImpactVector.from_dict(payload.get("use_heating", {})),
            use_water=ImpactVector.from_dict(payload.get("use_water", {})),
        )


@dataclass(slots=True)
class ConstructionItem:
    assembly: str
    material_type: str
    amount: float


@dataclass(slots=True)
class CogenerationInputs:
    fuel_type: str | None = None
    electricity_kwh: float = 0.0
    heating_mj: float = 0.0
    cooling_kwh: float = 0.0
    electricity_split: float = 0.0
    heating_split: float = 0.0
    cooling_split: float = 0.0


@dataclass(slots=True)
class WaterUseInputs:
    toilet_gpf: float = 0.0
    urinal_gpf: float = 0.0
    wc_sink_gpm: float = 0.0
    lab_sink_gpm: float = 0.0
    kitchen_sink_gpm: float = 0.0
    shower_gpm: float = 0.0
    landscaping_gal: float = 0.0
    rainwater_collection_gal: float = 0.0


@dataclass(slots=True)
class UsePhaseInputs:
    electricity_from_grid_kwh: float = 0.0
    onsite_renewable_kwh: float = 0.0
    natural_gas_m3: float = 0.0
    cogeneration: CogenerationInputs = field(default_factory=CogenerationInputs)
    water_use: WaterUseInputs = field(default_factory=WaterUseInputs)


@dataclass(slots=True)
class STVInputs:
    team: str
    construction_items: list[ConstructionItem]
    use_phase: UsePhaseInputs = field(default_factory=UsePhaseInputs)

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> "STVInputs":
        construction_items = [
            ConstructionItem(
                assembly=item["assembly"],
                material_type=item["material_type"],
                amount=float(item["amount"]),
            )
            for item in payload.get("construction_items", [])
        ]
        use_phase_payload = payload.get("use_phase", {})
        cogen_payload = use_phase_payload.get("cogeneration", {})
        water_payload = use_phase_payload.get("water_use", {})
        return cls(
            team=payload["team"],
            construction_items=construction_items,
            use_phase=UsePhaseInputs(
                electricity_from_grid_kwh=float(
                    use_phase_payload.get("electricity_from_grid_kwh", 0.0)
                ),
                onsite_renewable_kwh=float(
                    use_phase_payload.get("onsite_renewable_kwh", 0.0)
                ),
                natural_gas_m3=float(use_phase_payload.get("natural_gas_m3", 0.0)),
                cogeneration=CogenerationInputs(
                    fuel_type=cogen_payload.get("fuel_type"),
                    electricity_kwh=float(cogen_payload.get("electricity_kwh", 0.0)),
                    heating_mj=float(cogen_payload.get("heating_mj", 0.0)),
                    cooling_kwh=float(cogen_payload.get("cooling_kwh", 0.0)),
                    electricity_split=float(
                        cogen_payload.get("electricity_split", 0.0)
                    ),
                    heating_split=float(cogen_payload.get("heating_split", 0.0)),
                    cooling_split=float(cogen_payload.get("cooling_split", 0.0)),
                ),
                water_use=WaterUseInputs(
                    toilet_gpf=float(water_payload.get("toilet_gpf", 0.0)),
                    urinal_gpf=float(water_payload.get("urinal_gpf", 0.0)),
                    wc_sink_gpm=float(water_payload.get("wc_sink_gpm", 0.0)),
                    lab_sink_gpm=float(water_payload.get("lab_sink_gpm", 0.0)),
                    kitchen_sink_gpm=float(water_payload.get("kitchen_sink_gpm", 0.0)),
                    shower_gpm=float(water_payload.get("shower_gpm", 0.0)),
                    landscaping_gal=float(water_payload.get("landscaping_gal", 0.0)),
                    rainwater_collection_gal=float(
                        water_payload.get("rainwater_collection_gal", 0.0)
                    ),
                ),
            ),
        )


@dataclass(slots=True)
class ConstructionImpactResult:
    assembly: str
    material_type: str
    amount: float
    unit_multiplier: float
    embodied_total: ImpactVector
    materials: ImpactVector
    transport: ImpactVector
    construction: ImpactVector

    def to_dict(self) -> dict[str, Any]:
        return {
            "assembly": self.assembly,
            "material_type": self.material_type,
            "amount": self.amount,
            "unit_multiplier": self.unit_multiplier,
            "embodied_total": self.embodied_total.to_dict(),
            "materials": self.materials.to_dict(),
            "transport": self.transport.to_dict(),
            "construction": self.construction.to_dict(),
        }

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> "ConstructionImpactResult":
        return cls(
            assembly=payload["assembly"],
            material_type=payload["material_type"],
            amount=float(payload["amount"]),
            unit_multiplier=float(payload.get("unit_multiplier", 1.0)),
            embodied_total=ImpactVector.from_dict(payload.get("embodied_total", {})),
            materials=ImpactVector.from_dict(payload.get("materials", {})),
            transport=ImpactVector.from_dict(payload.get("transport", {})),
            construction=ImpactVector.from_dict(payload.get("construction", {})),
        )


@dataclass(slots=True)
class STVResults:
    team: str
    targets: ImpactVector
    breakdown: ImpactBreakdown
    construction_items: list[ConstructionImpactResult]
    lifetime_years: int

    def metric_summary(self) -> dict[str, dict[str, float | None]]:
        totals = self.breakdown.life_cycle
        summary: dict[str, dict[str, float | None]] = {}
        for key in IMPACT_KEYS:
            target = self.targets.get(key)
            project = totals.get(key)
            pct = None if target == 0 else project / target
            summary[key] = {"target": target, "project": project, "percent_of_target": pct}
        return summary

    def to_dict(self) -> dict[str, Any]:
        return {
            "team": self.team,
            "targets": self.targets.to_dict(),
            "metric_summary": self.metric_summary(),
            "breakdown": self.breakdown.to_dict(),
            "construction_items": [item.to_dict() for item in self.construction_items],
            "lifetime_years": self.lifetime_years,
        }

    def to_json_ready(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> "STVResults":
        return cls(
            team=payload["team"],
            targets=ImpactVector.from_dict(payload.get("targets", {})),
            breakdown=ImpactBreakdown.from_dict(payload.get("breakdown", {})),
            construction_items=[
                ConstructionImpactResult.from_dict(item)
                for item in payload.get("construction_items", [])
            ],
            lifetime_years=int(payload.get("lifetime_years", 0)),
        )

    @classmethod
    def combine(cls, results: list["STVResults"], *, team: str | None = None) -> "STVResults":
        if not results:
            raise ValueError("At least one STV result is required to create a project STV.")

        first = results[0]
        combined_team = team or first.team
        combined_targets = first.targets
        combined_lifetime = first.lifetime_years
        combined_breakdown = ImpactBreakdown()
        combined_items: list[ConstructionImpactResult] = []

        for result in results:
            if result.team != first.team:
                raise ValueError(
                    f"Cannot combine STV results from different teams: '{first.team}' and '{result.team}'."
                )
            if result.lifetime_years != combined_lifetime:
                raise ValueError(
                    "Cannot combine STV results with different lifetimes: "
                    f"{combined_lifetime} and {result.lifetime_years}."
                )
            if result.targets.to_dict() != combined_targets.to_dict():
                raise ValueError("Cannot combine STV results with different target values.")

            combined_breakdown = combined_breakdown + result.breakdown
            combined_items.extend(result.construction_items)

        return cls(
            team=combined_team,
            targets=combined_targets,
            breakdown=combined_breakdown,
            construction_items=combined_items,
            lifetime_years=combined_lifetime,
        )
