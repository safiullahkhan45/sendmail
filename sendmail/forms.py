from django import forms


class SendMailForm(forms.Form):

	credentials = forms.FileField()
	receivers = forms.FileField()
	subject = forms.CharField(max_length=200)
	body = forms.CharField(max_length=500)
	# attachment = forms.FileField()