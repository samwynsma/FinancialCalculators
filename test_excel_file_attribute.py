from apr_dpr_calc import AppreciateDepreciateCalculator
from budget_calc import BudgetMaker, budget_maker
from college_calc import CollegeSavingsCalculator
from excel_doc import ExcelDocument
from food_calc import FoodCostCalculator
from home_calc import HomeAffordabilityCalculator
from inherit_calc import InheritanceCalculator
from interest_calc import InterestCalculator
from loans_calc import LoanCalculator
from net_worth_calc import NetWorthCalculator
from retirement_calc import RetirementDurationCalculator
from retirement_goal_calc import RetirementCalculator
from take_home_calc import TakeHomeCalculator


CALCULATORS = [
    AppreciateDepreciateCalculator,
    BudgetMaker,
    CollegeSavingsCalculator,
    FoodCostCalculator,
    HomeAffordabilityCalculator,
    InheritanceCalculator,
    InterestCalculator,
    LoanCalculator,
    NetWorthCalculator,
    RetirementDurationCalculator,
    RetirementCalculator,
    TakeHomeCalculator,
]


def test_each_calculator_has_excel_file_path():
    for calculator_class in CALCULATORS:
        calculator = calculator_class()
        assert hasattr(calculator, "excel_file")
        assert calculator.excel_file == "InterestCalculation.xlsx"


def test_each_calculator_accepts_excel_document():
    document = ExcelDocument()
    for calculator_class in CALCULATORS:
        calculator = calculator_class(document=document)
        assert calculator.document is document


def test_budget_maker_accepts_excel_document(monkeypatch):
    document = ExcelDocument()
    captured = {}

    def fake_create_gui(self):
        captured["document"] = self.document
        captured["file_exists"] = self.file_exists

    monkeypatch.setattr(BudgetMaker, "create_gui", fake_create_gui)

    result = budget_maker(document)

    assert result is False
    assert captured["document"] is document
    assert captured["file_exists"] is False
