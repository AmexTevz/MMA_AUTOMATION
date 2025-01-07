import os
from dotenv import load_dotenv
from pathlib import Path


class Config:
    def __init__(self):
        # Load environment variables from .env file
        env_path = Path('.') / '.env'
        load_dotenv(dotenv_path=env_path)

        # Get credentials with fallbacks to None
        self.username = os.getenv('USERNAME')
        self.password = os.getenv('PASSWORD')
        self.app_path = os.getenv('APP_PATH')

    def validate(self):
        """Validate that all required credentials are present"""
        missing = []
        if not self.username:
            missing.append('USERNAME')
        if not self.password:
            missing.append('PASSWORD')
        if not self.app_path:
            missing.append('APP_PATH')
        
        if missing:
            raise ValueError(
                f"Missing required environment variables: {', '.join(missing)}. "
                "Please check your .env file."
            )

# Create a singleton instance
config = Config() 