from datetime import datetime
from bbgrl_slide_generator_v2_template import BBGRLSlideGeneratorV2

gen = BBGRLSlideGeneratorV2()
gen.generate_presentation(datetime(2025, 11, 10))
print('Test complete!')
