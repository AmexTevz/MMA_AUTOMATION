from dotenv import load_dotenv
import os

class Config:
    _instance = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            # Initialize the singleton instance
            load_dotenv()
            cls._instance.app_path = os.getenv('APP_PATH')
            cls._instance.domain = os.getenv('DOMAIN')
            cls._instance.user = os.getenv('USER')
            cls._instance.password = os.getenv('PASSWORD')
        return cls._instance
    
    def validate(self):
        required = ['app_path', 'domain', 'user', 'password']
        missing = [attr for attr in required if not getattr(self, attr)]
        if missing:
            raise ValueError(f"Missing required configuration: {', '.join(missing)}")

config = Config() 