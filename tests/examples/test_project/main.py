"""
Test project main entry point
"""

from services.api_client import get_data


def main():
    print("Starting test application...")
    data = get_data()
    print(f"Received data: {data}")
    print("Test application finished.")

if __name__ == "__main__":
    main() 

