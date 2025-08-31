# For ISD element type, do the following


- add element_type as ISD and it works only for stage 1
- after that in the optional object add an ISD object with list of element types you want to set. 
- make sure you add enough element types in that list which match the count of els_range value.
- add as many combinations you want to try

Sample config for ISD:

```
"optional": {
    "ISD": [
        ["NF90-4040", "NF90-4040", "NF90-4040", "NF90-4040", "NF200-4040", "NF200-4040"],
        ["NF90-4040", "NF90-4040", "NF90-4040", "NF200-4040", "NF200-4040", "NF200-4040"],
        ["NF90-4040", "NF90-4040", "NF200-4040", "NF200-4040", "NF200-4040", "NF200-4040"],
        ["NF90-4040", "NF90-4040", "NF90-4040", "NF90-4040", "NF200-4040", "NF200-4040"],
        ["NF90-4040", "NF90-4040", "NF90-4040", "NF200-4040", "NF200-4040", "NF200-4040"],
        ["NF90-4040", "NF90-4040", "NF200-4040", "NF200-4040", "NF200-4040", "NF200-4040"]
    ]
},
"stages": [
    {
        "stage_number": 1,
        "element_type": "NF90-4040",
        "pv_range": [1, 1],
        "els_range": [6, 6],
        "feed_pressure_range": [5, 6],
        "feed_pressure_step": 1
    },
```