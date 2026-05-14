#!/usr/bin/env python
"""Quick test to verify the psychrometric chart addition to thermal_comfort_ppt.py"""

import sys
import traceback

# Test: Import the module
try:
    from pages.modules import thermal_comfort_ppt
    print("✓ Successfully imported thermal_comfort_ppt module")
except Exception as e:
    print(f"✗ Failed to import thermal_comfort_ppt: {e}")
    traceback.print_exc()
    sys.exit(1)

# Test: Check if the new function exists
try:
    assert hasattr(thermal_comfort_ppt, 'plot_psychrometric_chart')
    print("✓ plot_psychrometric_chart function exists")
except AssertionError:
    print("✗ plot_psychrometric_chart function not found")
    sys.exit(1)

# Test: Check function signature
try:
    import inspect
    sig = inspect.signature(thermal_comfort_ppt.plot_psychrometric_chart)
    print(f"✓ Function signature: {sig}")
except Exception as e:
    print(f"✗ Failed to get function signature: {e}")
    sys.exit(1)

# Test: Create dummy data to verify function can run
try:
    import pandas as pd
    import numpy as np
    
    # Create test data
    dates = pd.date_range('2023-01-01', periods=8760, freq='H')
    test_df = pd.DataFrame({
        'datetime': dates,
        'dry_bulb_temperature': np.random.uniform(15, 35, 8760),
        'relative_humidity': np.random.uniform(30, 80, 8760),
    })
    
    # Call the function
    fig = thermal_comfort_ppt.plot_psychrometric_chart(test_df)
    print(f"✓ plot_psychrometric_chart executed successfully and returned {type(fig)}")
    
    # Close the figure to free memory
    import matplotlib.pyplot as plt
    plt.close(fig)
    
except Exception as e:
    print(f"✗ Failed to execute plot_psychrometric_chart: {e}")
    traceback.print_exc()
    sys.exit(1)

print("\n✓ All tests passed! The psychrometric chart has been successfully added to thermal_comfort_ppt.py")
