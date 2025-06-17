import pandas as pd
from typing import Dict, List
from src.models.data_models import AnticiposConfig

class AnticiposOnlineProcessor:
    def __init__(self, config: AnticiposConfig = None):
        self.config = config if config else AnticiposConfig()
    def validate_input_file(self, df: pd.DataFrame) -> bool:
       


