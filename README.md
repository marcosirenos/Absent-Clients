## Absent Clients

---

#### Objective

This is a project for automating the workflow of extracting data from flex monitor tool, process the download data, upload it to the database and refresh the Power Bi report


---

#### I've done it


### Reason

For better project structure it’s very useful to use files as packages, follow the example:

We want to create a calculator in Python but with operations outside the main, so can create a folder called “*operations*” and a main file called [calculator.py](http://calculator.py).:

```python
# inside operations folder we created a file called multiplication.py
class Multiplication:
	def multiply(self, a, b):
		return a * b
```

And in our main file:
