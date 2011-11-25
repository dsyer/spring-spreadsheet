{% capture sequence %}
<div class="wsd" wsd_style="qsd">
<pre>
{% endcapture %}

{% capture endsequence %}
</pre>
</div>
{% endcapture %}

{% comment %}

Example: usage:

{{ sequence }}
A->B: foo
B->A: bar
{{ endsequence }}

{% endcomment %}
