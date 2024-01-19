
## Summary

The `agenda.py` is a simple script which replaces the `${TAG}` with value in "Meeting Information" table on word document.

Beside that, there're 3 special tags: `${TB}`, `${TS}` and `${TSS}`, which are used for time information. `${TB}` is the beginning time of meeting, `${TS}` will sum up all minutes in `(xx')` on above session, and `${TSS}` will add one minute on `${TS}`. One exception is the first `${TS}` add 10 minutes on `${TB}`.

## Requirement

```bash
pip install python-docx
```
