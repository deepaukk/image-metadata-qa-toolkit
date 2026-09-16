# Delivery Packager

Builds a structured delivery folder from metadata + image/JSON pairs.

## Output layout

```
{output}/
  YYYYMMDD/
    {participant}/{locale}/{country}/{category}/
      image.jpg
      image.json
  YYYYMMDD_Metadata.xlsx    # outside the date folder
```

```bash
python delivery_packager_gui.py
```
