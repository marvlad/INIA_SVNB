pip install msoffcrypto-tool

python3 extract_colorimetric_mri_d_db.py \
  --input-dir "/Users/mascenci/Downloads/Clientes" \
  --sqlite "database_colorimetric/colorimetric_mri_d.sqlite" \
  --csv "database_colorimetric/colorimetric_mri_d.csv"




  Run examples:

python3 query_colorimetric_mri_d.py \
  --sqlite "database_colorimetric/colorimetric_mri_d.sqlite" \
  --tipo MRI \
  --month 5
python3 query_colorimetric_mri_d.py \
  --sqlite "database_colorimetric/colorimetric_mri_d.sqlite" \
  --tipo D \
  --month 2026-05

With source information for checking:

python3 query_colorimetric_mri_d.py \
  --sqlite "database_colorimetric/colorimetric_mri_d.sqlite" \
  --tipo D \
  --month 5 \
  --show-source

Only OLSEN:

python3 query_colorimetric_mri_d.py \
  --tipo MRI \
  --month 5 \
  --metodo OLSEN



