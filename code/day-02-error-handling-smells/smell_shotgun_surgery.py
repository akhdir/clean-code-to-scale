# adding a new "gift_wrap" order option touches all of these:
order_form.py       # add the checkbox
order_model.py      # add the field
pricing.py           # add the surcharge
invoice_pdf.py       # add the line item
shipping_label.py    # add the packing note

# one product decision, five unrelated files edited in
# lockstep - nothing about "gift wrap" lives in one place
