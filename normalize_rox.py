import argparse

from process_sybr_rox import run_rox_pipeline


### change Horizontal if you want A1, B1, C1, ... sorting instead of A1, A2, A3, ...
### change Vertical if you want A1, A2, A3, ... sorting instead of A1, B1, C1, ...

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Run the full SYBR/ROX preprocessing and normalization pipeline"
    )
    parser.add_argument(
        "--orientation",
        choices=["vertical", "horizontal"],
        default="vertical",  # right here you can change the default sorting orientation
        help="Well sorting orientation for preprocessing (default: vertical)",
    )
    parser.add_argument(
        "--minutes",
        type=int,
        default=5,  # change number of last time points used for ROX averaging
        help="Number of last time points used for ROX averaging (default: 5)",
    )
    args = parser.parse_args()


# input file name

    input_file = "Single Cas12 Pooled 4 GN gRNAs against GN with rox_20260429_231114_Multicomponent_20260504_155333.xlsx"
    
# experiment information for saving the final version
    experiment_info = {
        'Date': '20260429',
        'Type of Experiment': ['CRISPR single'],
        'Experiment Number': ['1'],
        'Pathogen': ['GN'],
        'Condition': ['pool 4 gRNAs -GN - test']
    }

    results = run_rox_pipeline(
        input_file=input_file,
        experiment_info=experiment_info,
        orientation=args.orientation,
        minutes=args.minutes,
    )

    print("Pipeline complete.")
    print(f"Output file: {results['output_file']}")
