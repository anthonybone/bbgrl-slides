from datetime import datetime


def main():
	from bbgrl_slide_generator_v1 import bbgrlslidegeneratorv1
	gen = bbgrlslidegeneratorv1()
	gen.generate_presentation(datetime(2025, 11, 10))
	print('Test complete!')


def test_integration_manual():
	import pytest
	pytest.skip("Integration script; run directly: python test/test_heart_of_jesus.py")


if __name__ == "__main__":
	main()
