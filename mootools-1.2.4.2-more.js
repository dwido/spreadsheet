//MooTools More, <http://mootools.net/more>. Copyright (c) 2006-2009 Aaron Newton <http://clientcide.com/>, Valerio Proietti <http://mad4milk.net> & the MooTools team <http://mootools.net/developers>, MIT Style License.

/*
---

script: More.js

description: MooTools More

license: MIT-style license

authors:
- Guillermo Rauch
- Thomas Aylott
- Scott Kyle

requires:
- core:1.2.4/MooTools

provides: [MooTools.More]

...
*/

MooTools.More = {
	'version': '1.2.4.2',
	'build': 'bd5a93c0913cce25917c48cbdacde568e15e02ef'
};

/*
---

script: Class.Binds.js

description: Automagically binds specified methods in a class to the instance of the class.

license: MIT-style license

authors:
- Aaron Newton

requires:
- core:1.2.4/Class
- /MooTools.More

provides: [Class.Binds]

...
*/

Class.Mutators.Binds = function(binds){
    return binds;
};

Class.Mutators.initialize = function(initialize){
	return function(){
		$splat(this.Binds).each(function(name){
			var original = this[name];
			if (original) this[name] = original.bind(this);
		}, this);
		return initialize.apply(this, arguments);
	};
};


/*
---

script: Element.Forms.js

description: Extends the Element native object to include methods useful in managing inputs.

license: MIT-style license

authors:
- Aaron Newton

requires:
- core:1.2.4/Element
- /MooTools.More

provides: [Element.Forms]

...
*/

Element.implement({

	tidy: function(){
		this.set('value', this.get('value').tidy());
	},

	getTextInRange: function(start, end){
		return this.get('value').substring(start, end);
	},

	getSelectedText: function(){
		if (this.setSelectionRange) return this.getTextInRange(this.getSelectionStart(), this.getSelectionEnd());
		return document.selection.createRange().text;
	},

	getSelectedRange: function() {
		if ($defined(this.selectionStart)) return {start: this.selectionStart, end: this.selectionEnd};
		var pos = {start: 0, end: 0};
		var range = this.getDocument().selection.createRange();
		if (!range || range.parentElement() != this) return pos;
		var dup = range.duplicate();
		if (this.type == 'text') {
			pos.start = 0 - dup.moveStart('character', -100000);
			pos.end = pos.start + range.text.length;
		} else {
			var value = this.get('value');
			var offset = value.length;
			dup.moveToElementText(this);
			dup.setEndPoint('StartToEnd', range);
			if(dup.text.length) offset -= value.match(/[\n\r]*$/)[0].length;
			pos.end = offset - dup.text.length;
			dup.setEndPoint('StartToStart', range);
			pos.start = offset - dup.text.length;
		}
		return pos;
	},

	getSelectionStart: function(){
		return this.getSelectedRange().start;
	},

	getSelectionEnd: function(){
		return this.getSelectedRange().end;
	},

	setCaretPosition: function(pos){
		if (pos == 'end') pos = this.get('value').length;
		this.selectRange(pos, pos);
		return this;
	},

	getCaretPosition: function(){
		return this.getSelectedRange().start;
	},

	selectRange: function(start, end){
		if (this.setSelectionRange) {
			this.focus();
			this.setSelectionRange(start, end);
		} else {
			var value = this.get('value');
			var diff = value.substr(start, end - start).replace(/\r/g, '').length;
			start = value.substr(0, start).replace(/\r/g, '').length;
			var range = this.createTextRange();
			range.collapse(true);
			range.moveEnd('character', start + diff);
			range.moveStart('character', start);
			range.select();
		}
		return this;
	},

	insertAtCursor: function(value, select){
		var pos = this.getSelectedRange();
		var text = this.get('value');
		this.set('value', text.substring(0, pos.start) + value + text.substring(pos.end, text.length));
		if ($pick(select, true)) this.selectRange(pos.start, pos.start + value.length);
		else this.setCaretPosition(pos.start + value.length);
		return this;
	},

	insertAroundCursor: function(options, select){
		options = $extend({
			before: '',
			defaultMiddle: '',
			after: ''
		}, options);
		var value = this.getSelectedText() || options.defaultMiddle;
		var pos = this.getSelectedRange();
		var text = this.get('value');
		if (pos.start == pos.end){
			this.set('value', text.substring(0, pos.start) + options.before + value + options.after + text.substring(pos.end, text.length));
			this.selectRange(pos.start + options.before.length, pos.end + options.before.length + value.length);
		} else {
			var current = text.substring(pos.start, pos.end);
			this.set('value', text.substring(0, pos.start) + options.before + current + options.after + text.substring(pos.end, text.length));
			var selStart = pos.start + options.before.length;
			if ($pick(select, true)) this.selectRange(selStart, selStart + current.length);
			else this.setCaretPosition(selStart + text.length);
		}
		return this;
	}

});

/*
---

script: Element.Measure.js

description: Extends the Element native object to include methods useful in measuring dimensions.

credits: "Element.measure / .expose methods by Daniel Steigerwald License: MIT-style license. Copyright: Copyright (c) 2008 Daniel Steigerwald, daniel.steigerwald.cz"

license: MIT-style license

authors:
- Aaron Newton

requires:
- core:1.2.4/Element.Style
- core:1.2.4/Element.Dimensions
- /MooTools.More

provides: [Element.Measure]

...
*/

Element.implement({

	measure: function(fn){
		var vis = function(el) {
			return !!(!el || el.offsetHeight || el.offsetWidth);
		};
		if (vis(this)) return fn.apply(this);
		var parent = this.getParent(),
			restorers = [],
			toMeasure = []; 
		while (!vis(parent) && parent != document.body) {
			toMeasure.push(parent.expose());
			parent = parent.getParent();
		}
		var restore = this.expose();
		var result = fn.apply(this);
		restore();
		toMeasure.each(function(restore){
			restore();
		});
		return result;
	},

	expose: function(){
		if (this.getStyle('display') != 'none') return $empty;
		var before = this.style.cssText;
		this.setStyles({
			display: 'block',
			position: 'absolute',
			visibility: 'hidden'
		});
		return function(){
			this.style.cssText = before;
		}.bind(this);
	},

	getDimensions: function(options){
		options = $merge({computeSize: false},options);
		var dim = {};
		var getSize = function(el, options){
			return (options.computeSize)?el.getComputedSize(options):el.getSize();
		};
		var parent = this.getParent('body');
		if (parent && this.getStyle('display') == 'none'){
			dim = this.measure(function(){
				return getSize(this, options);
			});
		} else if (parent){
			try { //safari sometimes crashes here, so catch it
				dim = getSize(this, options);
			}catch(e){}
		} else {
			dim = {x: 0, y: 0};
		}
		return $chk(dim.x) ? $extend(dim, {width: dim.x, height: dim.y}) : $extend(dim, {x: dim.width, y: dim.height});
	},

	getComputedSize: function(options){
		options = $merge({
			styles: ['padding','border'],
			plains: {
				height: ['top','bottom'],
				width: ['left','right']
			},
			mode: 'both'
		}, options);
		var size = {width: 0,height: 0};
		switch (options.mode){
			case 'vertical':
				delete size.width;
				delete options.plains.width;
				break;
			case 'horizontal':
				delete size.height;
				delete options.plains.height;
				break;
		}
		var getStyles = [];
		//this function might be useful in other places; perhaps it should be outside this function?
		$each(options.plains, function(plain, key){
			plain.each(function(edge){
				options.styles.each(function(style){
					getStyles.push((style == 'border') ? style + '-' + edge + '-' + 'width' : style + '-' + edge);
				});
			});
		});
		var styles = {};
		getStyles.each(function(style){ styles[style] = this.getComputedStyle(style); }, this);
		var subtracted = [];
		$each(options.plains, function(plain, key){ //keys: width, height, plains: ['left', 'right'], ['top','bottom']
			var capitalized = key.capitalize();
			size['total' + capitalized] = size['computed' + capitalized] = 0;
			plain.each(function(edge){ //top, left, right, bottom
				size['computed' + edge.capitalize()] = 0;
				getStyles.each(function(style, i){ //padding, border, etc.
					//'padding-left'.test('left') size['totalWidth'] = size['width'] + [padding-left]
					if (style.test(edge)){
						styles[style] = styles[style].toInt() || 0; //styles['padding-left'] = 5;
						size['total' + capitalized] = size['total' + capitalized] + styles[style];
						size['computed' + edge.capitalize()] = size['computed' + edge.capitalize()] + styles[style];
					}
					//if width != width (so, padding-left, for instance), then subtract that from the total
					if (style.test(edge) && key != style &&
						(style.test('border') || style.test('padding')) && !subtracted.contains(style)){
						subtracted.push(style);
						size['computed' + capitalized] = size['computed' + capitalized]-styles[style];
					}
				});
			});
		});

		['Width', 'Height'].each(function(value){
			var lower = value.toLowerCase();
			if(!$chk(size[lower])) return;

			size[lower] = size[lower] + this['offset' + value] + size['computed' + value];
			size['total' + value] = size[lower] + size['total' + value];
			delete size['computed' + value];
		}, this);

		return $extend(styles, size);
	}

});

/*
---

script: Element.Position.js

description: Extends the Element native object to include methods useful positioning elements relative to others.

license: MIT-style license

authors:
- Aaron Newton

requires:
- core:1.2.4/Element.Dimensions
- /Element.Measure

provides: [Elements.Position]

...
*/

(function(){

var original = Element.prototype.position;

Element.implement({

	position: function(options){
		//call original position if the options are x/y values
		if (options && ($defined(options.x) || $defined(options.y))) return original ? original.apply(this, arguments) : this;
		$each(options||{}, function(v, k){ if (!$defined(v)) delete options[k]; });
		options = $merge({
			// minimum: { x: 0, y: 0 },
			// maximum: { x: 0, y: 0},
			relativeTo: document.body,
			position: {
				x: 'center', //left, center, right
				y: 'center' //top, center, bottom
			},
			edge: false,
			offset: {x: 0, y: 0},
			returnPos: false,
			relFixedPosition: false,
			ignoreMargins: false,
			ignoreScroll: false,
			allowNegative: false
		}, options);
		//compute the offset of the parent positioned element if this element is in one
		var parentOffset = {x: 0, y: 0}, 
				parentPositioned = false;
		/* dollar around getOffsetParent should not be necessary, but as it does not return
		 * a mootools extended element in IE, an error occurs on the call to expose. See:
		 * http://mootools.lighthouseapp.com/projects/2706/tickets/333-element-getoffsetparent-inconsistency-between-ie-and-other-browsers */
		var offsetParent = this.measure(function(){
			return document.id(this.getOffsetParent());
		});
		if (offsetParent && offsetParent != this.getDocument().body){
			parentOffset = offsetParent.measure(function(){
				return this.getPosition();
			});
			parentPositioned = offsetParent != document.id(options.relativeTo);
			options.offset.x = options.offset.x - parentOffset.x;
			options.offset.y = options.offset.y - parentOffset.y;
		}
		//upperRight, bottomRight, centerRight, upperLeft, bottomLeft, centerLeft
		//topRight, topLeft, centerTop, centerBottom, center
		var fixValue = function(option){
			if ($type(option) != 'string') return option;
			option = option.toLowerCase();
			var val = {};
			if (option.test('left')) val.x = 'left';
			else if (option.test('right')) val.x = 'right';
			else val.x = 'center';
			if (option.test('upper') || option.test('top')) val.y = 'top';
			else if (option.test('bottom')) val.y = 'bottom';
			else val.y = 'center';
			return val;
		};
		options.edge = fixValue(options.edge);
		options.position = fixValue(options.position);
		if (!options.edge){
			if (options.position.x == 'center' && options.position.y == 'center') options.edge = {x:'center', y:'center'};
			else options.edge = {x:'left', y:'top'};
		}

		this.setStyle('position', 'absolute');
		var rel = document.id(options.relativeTo) || document.body,
				calc = rel == document.body ? window.getScroll() : rel.getPosition(),
				top = calc.y, left = calc.x;

		var scrolls = rel.getScrolls();
		top += scrolls.y;
		left += scrolls.x;

		var dim = this.getDimensions({computeSize: true, styles:['padding', 'border','margin']});
		var pos = {},
				prefY = options.offset.y,
				prefX = options.offset.x,
				winSize = window.getSize();
		switch(options.position.x){
			case 'left':
				pos.x = left + prefX;
				break;
			case 'right':
				pos.x = left + prefX + rel.offsetWidth;
				break;
			default: //center
				pos.x = left + ((rel == document.body ? winSize.x : rel.offsetWidth)/2) + prefX;
				break;
		}
		switch(options.position.y){
			case 'top':
				pos.y = top + prefY;
				break;
			case 'bottom':
				pos.y = top + prefY + rel.offsetHeight;
				break;
			default: //center
				pos.y = top + ((rel == document.body ? winSize.y : rel.offsetHeight)/2) + prefY;
				break;
		}
		if (options.edge){
			var edgeOffset = {};

			switch(options.edge.x){
				case 'left':
					edgeOffset.x = 0;
					break;
				case 'right':
					edgeOffset.x = -dim.x-dim.computedRight-dim.computedLeft;
					break;
				default: //center
					edgeOffset.x = -(dim.totalWidth/2);
					break;
			}
			switch(options.edge.y){
				case 'top':
					edgeOffset.y = 0;
					break;
				case 'bottom':
					edgeOffset.y = -dim.y-dim.computedTop-dim.computedBottom;
					break;
				default: //center
					edgeOffset.y = -(dim.totalHeight/2);
					break;
			}
			pos.x += edgeOffset.x;
			pos.y += edgeOffset.y;
		}
		pos = {
			left: ((pos.x >= 0 || parentPositioned || options.allowNegative) ? pos.x : 0).toInt(),
			top: ((pos.y >= 0 || parentPositioned || options.allowNegative) ? pos.y : 0).toInt()
		};
		var xy = {left: 'x', top: 'y'};
		['minimum', 'maximum'].each(function(minmax) {
			['left', 'top'].each(function(lr) {
				var val = options[minmax] ? options[minmax][xy[lr]] : null;
				if (val != null && pos[lr] < val) pos[lr] = val;
			});
		});
		if (rel.getStyle('position') == 'fixed' || options.relFixedPosition){
			var winScroll = window.getScroll();
			pos.top+= winScroll.y;
			pos.left+= winScroll.x;
		}
		if (options.ignoreScroll) {
			var relScroll = rel.getScroll();
			pos.top-= relScroll.y;
			pos.left-= relScroll.x;
		}
		if (options.ignoreMargins) {
			pos.left += (
				options.edge.x == 'right' ? dim['margin-right'] : 
				options.edge.x == 'center' ? -dim['margin-left'] + ((dim['margin-right'] + dim['margin-left'])/2) : 
					- dim['margin-left']
			);
			pos.top += (
				options.edge.y == 'bottom' ? dim['margin-bottom'] : 
				options.edge.y == 'center' ? -dim['margin-top'] + ((dim['margin-bottom'] + dim['margin-top'])/2) : 
					- dim['margin-top']
			);
		}
		pos.left = Math.ceil(pos.left);
		pos.top = Math.ceil(pos.top);
		if (options.returnPos) return pos;
		else this.setStyles(pos);
		return this;
	}

});

})();

/*
---

script: Element.Shortcuts.js

description: Extends the Element native object to include some shortcut methods.

license: MIT-style license

authors:
- Aaron Newton

requires:
- core:1.2.4/Element.Style
- /MooTools.More

provides: [Element.Shortcuts]

...
*/

Element.implement({

	isDisplayed: function(){
		return this.getStyle('display') != 'none';
	},

	isVisible: function(){
		var w = this.offsetWidth,
			h = this.offsetHeight;
		return (w == 0 && h == 0) ? false : (w > 0 && h > 0) ? true : this.isDisplayed();
	},

	toggle: function(){
		return this[this.isDisplayed() ? 'hide' : 'show']();
	},

	hide: function(){
		var d;
		try {
			// IE fails here if the element is not in the dom
			if ((d = this.getStyle('display')) == 'none') d = null;
		} catch(e){}
		
		return this.store('originalDisplay', d || 'block').setStyle('display', 'none');
	},

	show: function(display){
		return this.setStyle('display', display || this.retrieve('originalDisplay') || 'block');
	},

	swapClass: function(remove, add){
		return this.removeClass(remove).addClass(add);
	}

});


/*
---

script: Fx.Elements.js

description: Effect to change any number of CSS properties of any number of Elements.

license: MIT-style license

authors:
- Valerio Proietti

requires:
- core:1.2.4/Fx.CSS
- /MooTools.More

provides: [Fx.Elements]

...
*/

Fx.Elements = new Class({

	Extends: Fx.CSS,

	initialize: function(elements, options){
		this.elements = this.subject = $$(elements);
		this.parent(options);
	},

	compute: function(from, to, delta){
		var now = {};
		for (var i in from){
			var iFrom = from[i], iTo = to[i], iNow = now[i] = {};
			for (var p in iFrom) iNow[p] = this.parent(iFrom[p], iTo[p], delta);
		}
		return now;
	},

	set: function(now){
		for (var i in now){
			var iNow = now[i];
			for (var p in iNow) this.render(this.elements[i], p, iNow[p], this.options.unit);
		}
		return this;
	},

	start: function(obj){
		if (!this.check(obj)) return this;
		var from = {}, to = {};
		for (var i in obj){
			var iProps = obj[i], iFrom = from[i] = {}, iTo = to[i] = {};
			for (var p in iProps){
				var parsed = this.prepare(this.elements[i], p, iProps[p]);
				iFrom[p] = parsed.from;
				iTo[p] = parsed.to;
			}
		}
		return this.parent(from, to);
	}

});

/*
---

script: Fx.Accordion.js

description: An Fx.Elements extension which allows you to easily create accordion type controls.

license: MIT-style license

authors:
- Valerio Proietti

requires:
- core:1.2.4/Element.Event
- /Fx.Elements

provides: [Fx.Accordion]

...
*/

var Accordion = Fx.Accordion = new Class({

	Extends: Fx.Elements,

	options: {/*
		onActive: $empty(toggler, section),
		onBackground: $empty(toggler, section),
		fixedHeight: false,
		fixedWidth: false,
		*/
		display: 0,
		show: false,
		height: true,
		width: false,
		opacity: true,
		alwaysHide: false,
		trigger: 'click',
		initialDisplayFx: true,
		returnHeightToAuto: true
	},

	initialize: function(){
		var params = Array.link(arguments, {'container': Element.type, 'options': Object.type, 'togglers': $defined, 'elements': $defined});
		this.parent(params.elements, params.options);
		this.togglers = $$(params.togglers);
		this.container = document.id(params.container);
		this.previous = -1;
		this.internalChain = new Chain();
		if (this.options.alwaysHide) this.options.wait = true;
		if ($chk(this.options.show)){
			this.options.display = false;
			this.previous = this.options.show;
		}
		if (this.options.start){
			this.options.display = false;
			this.options.show = false;
		}
		this.effects = {};
		if (this.options.opacity) this.effects.opacity = 'fullOpacity';
		if (this.options.width) this.effects.width = this.options.fixedWidth ? 'fullWidth' : 'offsetWidth';
		if (this.options.height) this.effects.height = this.options.fixedHeight ? 'fullHeight' : 'scrollHeight';
		for (var i = 0, l = this.togglers.length; i < l; i++) this.addSection(this.togglers[i], this.elements[i]);
		this.elements.each(function(el, i){
			if (this.options.show === i){
				this.fireEvent('active', [this.togglers[i], el]);
			} else {
				for (var fx in this.effects) el.setStyle(fx, 0);
			}
		}, this);
		if ($chk(this.options.display)) this.display(this.options.display, this.options.initialDisplayFx);
		this.addEvent('complete', this.internalChain.callChain.bind(this.internalChain));
	},

	addSection: function(toggler, element){
		toggler = document.id(toggler);
		element = document.id(element);
		var test = this.togglers.contains(toggler);
		this.togglers.include(toggler);
		this.elements.include(element);
		var idx = this.togglers.indexOf(toggler);
		var displayer = this.display.bind(this, idx);
		toggler.store('accordion:display', displayer);
		toggler.addEvent(this.options.trigger, displayer);
		if (this.options.height) element.setStyles({'padding-top': 0, 'border-top': 'none', 'padding-bottom': 0, 'border-bottom': 'none'});
		if (this.options.width) element.setStyles({'padding-left': 0, 'border-left': 'none', 'padding-right': 0, 'border-right': 'none'});
		element.fullOpacity = 1;
		if (this.options.fixedWidth) element.fullWidth = this.options.fixedWidth;
		if (this.options.fixedHeight) element.fullHeight = this.options.fixedHeight;
		element.setStyle('overflow', 'hidden');
		if (!test){
			for (var fx in this.effects) element.setStyle(fx, 0);
		}
		return this;
	},

	detach: function(){
		this.togglers.each(function(toggler) {
			toggler.removeEvent(this.options.trigger, toggler.retrieve('accordion:display'));
		}, this);
	},

	display: function(index, useFx){
		if (!this.check(index, useFx)) return this;
		useFx = $pick(useFx, true);
		if (this.options.returnHeightToAuto){
			var prev = this.elements[this.previous];
			if (prev && !this.selfHidden){
				for (var fx in this.effects){
					prev.setStyle(fx, prev[this.effects[fx]]);
				}
			}
		}
		index = ($type(index) == 'element') ? this.elements.indexOf(index) : index;
		if ((this.timer && this.options.wait) || (index === this.previous && !this.options.alwaysHide)) return this;
		this.previous = index;
		var obj = {};
		this.elements.each(function(el, i){
			obj[i] = {};
			var hide;
			if (i != index){
				hide = true;
			} else if (this.options.alwaysHide && ((el.offsetHeight > 0 && this.options.height) || el.offsetWidth > 0 && this.options.width)){
				hide = true;
				this.selfHidden = true;
			}
			this.fireEvent(hide ? 'background' : 'active', [this.togglers[i], el]);
			for (var fx in this.effects) obj[i][fx] = hide ? 0 : el[this.effects[fx]];
		}, this);
		this.internalChain.chain(function(){
			if (this.options.returnHeightToAuto && !this.selfHidden){
				var el = this.elements[index];
				if (el) el.setStyle('height', 'auto');
			};
		}.bind(this));
		return useFx ? this.start(obj) : this.set(obj);
	}

});

/*
---

script: Fx.Move.js

description: Defines Fx.Move, a class that works with Element.Position.js to transition an element from one location to another.

license: MIT-style license

authors:
- Aaron Newton

requires:
- core:1.2.4/Fx.Morph
- /Element.Position

provides: [Fx.Move]

...
*/

Fx.Move = new Class({

	Extends: Fx.Morph,

	options: {
		relativeTo: document.body,
		position: 'center',
		edge: false,
		offset: {x: 0, y: 0}
	},

	start: function(destination){
		return this.parent(this.element.position($merge(this.options, destination, {returnPos: true})));
	}

});

Element.Properties.move = {

	set: function(options){
		var morph = this.retrieve('move');
		if (morph) morph.cancel();
		return this.eliminate('move').store('move:options', $extend({link: 'cancel'}, options));
	},

	get: function(options){
		if (options || !this.retrieve('move')){
			if (options || !this.retrieve('move:options')) this.set('move', options);
			this.store('move', new Fx.Move(this, this.retrieve('move:options')));
		}
		return this.retrieve('move');
	}

};

Element.implement({

	move: function(options){
		this.get('move').start(options);
		return this;
	}

});


/*
---

script: Fx.Reveal.js

description: Defines Fx.Reveal, a class that shows and hides elements with a transition.

license: MIT-style license

authors:
- Aaron Newton

requires:
- core:1.2.4/Fx.Morph
- /Element.Shortcuts
- /Element.Measure

provides: [Fx.Reveal]

...
*/

Fx.Reveal = new Class({

	Extends: Fx.Morph,

	options: {/*	  
		onShow: $empty(thisElement),
		onHide: $empty(thisElement),
		onComplete: $empty(thisElement),
		heightOverride: null,
		widthOverride: null, */
		link: 'cancel',
		styles: ['padding', 'border', 'margin'],
		transitionOpacity: !Browser.Engine.trident4,
		mode: 'vertical',
		display: 'block',
		hideInputs: Browser.Engine.trident ? 'select, input, textarea, object, embed' : false
	},

	dissolve: function(){
		try {
			if (!this.hiding && !this.showing){
				if (this.element.getStyle('display') != 'none'){
					this.hiding = true;
					this.showing = false;
					this.hidden = true;
					this.cssText = this.element.style.cssText;
					var startStyles = this.element.getComputedSize({
						styles: this.options.styles,
						mode: this.options.mode
					});
					this.element.setStyle('display', 'block');
					if (this.options.transitionOpacity) startStyles.opacity = 1;
					var zero = {};
					$each(startStyles, function(style, name){
						zero[name] = [style, 0];
					}, this);
					this.element.setStyle('overflow', 'hidden');
					var hideThese = this.options.hideInputs ? this.element.getElements(this.options.hideInputs) : null;
					this.$chain.unshift(function(){
						if (this.hidden){
							this.hiding = false;
							$each(startStyles, function(style, name){
								startStyles[name] = style;
							}, this);
							this.element.style.cssText = this.cssText;
							this.element.setStyle('display', 'none');
							if (hideThese) hideThese.setStyle('visibility', 'visible');
						}
						this.fireEvent('hide', this.element);
						this.callChain();
					}.bind(this));
					if (hideThese) hideThese.setStyle('visibility', 'hidden');
					this.start(zero);
				} else {
					this.callChain.delay(10, this);
					this.fireEvent('complete', this.element);
					this.fireEvent('hide', this.element);
				}
			} else if (this.options.link == 'chain'){
				this.chain(this.dissolve.bind(this));
			} else if (this.options.link == 'cancel' && !this.hiding){
				this.cancel();
				this.dissolve();
			}
		} catch(e){
			this.hiding = false;
			this.element.setStyle('display', 'none');
			this.callChain.delay(10, this);
			this.fireEvent('complete', this.element);
			this.fireEvent('hide', this.element);
		}
		return this;
	},

	reveal: function(){
		try {
			if (!this.showing && !this.hiding){
				if (this.element.getStyle('display') == 'none' ||
					 this.element.getStyle('visiblity') == 'hidden' ||
					 this.element.getStyle('opacity') == 0){
					this.showing = true;
					this.hiding = this.hidden =  false;
					var startStyles;
					this.cssText = this.element.style.cssText;
					//toggle display, but hide it
					this.element.measure(function(){
						//create the styles for the opened/visible state
						startStyles = this.element.getComputedSize({
							styles: this.options.styles,
							mode: this.options.mode
						});
					}.bind(this));
					$each(startStyles, function(style, name){
						startStyles[name] = style;
					});
					//if we're overridding height/width
					if ($chk(this.options.heightOverride)) startStyles.height = this.options.heightOverride.toInt();
					if ($chk(this.options.widthOverride)) startStyles.width = this.options.widthOverride.toInt();
					if (this.options.transitionOpacity) {
						this.element.setStyle('opacity', 0);
						startStyles.opacity = 1;
					}
					//create the zero state for the beginning of the transition
					var zero = {
						height: 0,
						display: this.options.display
					};
					$each(startStyles, function(style, name){ zero[name] = 0; });
					//set to zero
					this.element.setStyles($merge(zero, {overflow: 'hidden'}));
					//hide inputs
					var hideThese = this.options.hideInputs ? this.element.getElements(this.options.hideInputs) : null;
					if (hideThese) hideThese.setStyle('visibility', 'hidden');
					//start the effect
					this.start(startStyles);
					this.$chain.unshift(function(){
						this.element.style.cssText = this.cssText;
						this.element.setStyle('display', this.options.display);
						if (!this.hidden) this.showing = false;
						if (hideThese) hideThese.setStyle('visibility', 'visible');
						this.callChain();
						this.fireEvent('show', this.element);
					}.bind(this));
				} else {
					this.callChain();
					this.fireEvent('complete', this.element);
					this.fireEvent('show', this.element);
				}
			} else if (this.options.link == 'chain'){
				this.chain(this.reveal.bind(this));
			} else if (this.options.link == 'cancel' && !this.showing){
				this.cancel();
				this.reveal();
			}
		} catch(e){
			this.element.setStyles({
				display: this.options.display,
				visiblity: 'visible',
				opacity: 1
			});
			this.showing = false;
			this.callChain.delay(10, this);
			this.fireEvent('complete', this.element);
			this.fireEvent('show', this.element);
		}
		return this;
	},

	toggle: function(){
		if (this.element.getStyle('display') == 'none' ||
			 this.element.getStyle('visiblity') == 'hidden' ||
			 this.element.getStyle('opacity') == 0){
			this.reveal();
		} else {
			this.dissolve();
		}
		return this;
	},

	cancel: function(){
		this.parent.apply(this, arguments);
		this.element.style.cssText = this.cssText;
		this.hidding = false;
		this.showing = false;
	}

});

Element.Properties.reveal = {

	set: function(options){
		var reveal = this.retrieve('reveal');
		if (reveal) reveal.cancel();
		return this.eliminate('reveal').store('reveal:options', options);
	},

	get: function(options){
		if (options || !this.retrieve('reveal')){
			if (options || !this.retrieve('reveal:options')) this.set('reveal', options);
			this.store('reveal', new Fx.Reveal(this, this.retrieve('reveal:options')));
		}
		return this.retrieve('reveal');
	}

};

Element.Properties.dissolve = Element.Properties.reveal;

Element.implement({

	reveal: function(options){
		this.get('reveal', options).reveal();
		return this;
	},

	dissolve: function(options){
		this.get('reveal', options).dissolve();
		return this;
	},

	nix: function(){
		var params = Array.link(arguments, {destroy: Boolean.type, options: Object.type});
		this.get('reveal', params.options).dissolve().chain(function(){
			this[params.destroy ? 'destroy' : 'dispose']();
		}.bind(this));
		return this;
	},

	wink: function(){
		var params = Array.link(arguments, {duration: Number.type, options: Object.type});
		var reveal = this.get('reveal', params.options);
		reveal.reveal().chain(function(){
			(function(){
				reveal.dissolve();
			}).delay(params.duration || 2000);
		});
	}


});

/*
---

script: Fx.Scroll.js

description: Effect to smoothly scroll any element, including the window.

license: MIT-style license

authors:
- Valerio Proietti

requires:
- core:1.2.4/Fx
- core:1.2.4/Element.Event
- core:1.2.4/Element.Dimensions
- /MooTools.More

provides: [Fx.Scroll]

...
*/

Fx.Scroll = new Class({

	Extends: Fx,

	options: {
		offset: {x: 0, y: 0},
		wheelStops: true
	},

	initialize: function(element, options){
		this.element = this.subject = document.id(element);
		this.parent(options);
		var cancel = this.cancel.bind(this, false);

		if ($type(this.element) != 'element') this.element = document.id(this.element.getDocument().body);

		var stopper = this.element;

		if (this.options.wheelStops){
			this.addEvent('start', function(){
				stopper.addEvent('mousewheel', cancel);
			}, true);
			this.addEvent('complete', function(){
				stopper.removeEvent('mousewheel', cancel);
			}, true);
		}
	},

	set: function(){
		var now = Array.flatten(arguments);
		if (Browser.Engine.gecko) now = [Math.round(now[0]), Math.round(now[1])];
		this.element.scrollTo(now[0], now[1]);
	},

	compute: function(from, to, delta){
		return [0, 1].map(function(i){
			return Fx.compute(from[i], to[i], delta);
		});
	},

	start: function(x, y){
		if (!this.check(x, y)) return this;
		var scrollSize = this.element.getScrollSize(),
			scroll = this.element.getScroll(), 
			values = {x: x, y: y};
		for (var z in values){
			var max = scrollSize[z];
			if ($chk(values[z])) values[z] = ($type(values[z]) == 'number') ? values[z] : max;
			else values[z] = scroll[z];
			values[z] += this.options.offset[z];
		}
		return this.parent([scroll.x, scroll.y], [values.x, values.y]);
	},

	toTop: function(){
		return this.start(false, 0);
	},

	toLeft: function(){
		return this.start(0, false);
	},

	toRight: function(){
		return this.start('right', false);
	},

	toBottom: function(){
		return this.start(false, 'bottom');
	},

	toElement: function(el){
		var position = document.id(el).getPosition(this.element);
		return this.start(position.x, position.y);
	},

	scrollIntoView: function(el, axes, offset){
		axes = axes ? $splat(axes) : ['x','y'];
		var to = {};
		el = document.id(el);
		var pos = el.getPosition(this.element);
		var size = el.getSize();
		var scroll = this.element.getScroll();
		var containerSize = this.element.getSize();
		var edge = {
			x: pos.x + size.x,
			y: pos.y + size.y
		};
		['x','y'].each(function(axis) {
			if (axes.contains(axis)) {
				if (edge[axis] > scroll[axis] + containerSize[axis]) to[axis] = edge[axis] - containerSize[axis];
				if (pos[axis] < scroll[axis]) to[axis] = pos[axis];
			}
			if (to[axis] == null) to[axis] = scroll[axis];
			if (offset && offset[axis]) to[axis] = to[axis] + offset[axis];
		}, this);
		if (to.x != scroll.x || to.y != scroll.y) this.start(to.x, to.y);
		return this;
	},

	scrollToCenter: function(el, axes, offset){
		axes = axes ? $splat(axes) : ['x', 'y'];
		el = $(el);
		var to = {},
			pos = el.getPosition(this.element),
			size = el.getSize(),
			scroll = this.element.getScroll(),
			containerSize = this.element.getSize(),
			edge = {
				x: pos.x + size.x,
				y: pos.y + size.y
			};

		['x','y'].each(function(axis){
			if(axes.contains(axis)){
				to[axis] = pos[axis] - (containerSize[axis] - size[axis])/2;
			}
			if(to[axis] == null) to[axis] = scroll[axis];
			if(offset && offset[axis]) to[axis] = to[axis] + offset[axis];
		}, this);
		if (to.x != scroll.x || to.y != scroll.y) this.start(to.x, to.y);
		return this;
	}

});


/*
---

script: Fx.Slide.js

description: Effect to slide an element in and out of view.

license: MIT-style license

authors:
- Valerio Proietti

requires:
- core:1.2.4/Fx Element.Style
- /MooTools.More

provides: [Fx.Slide]

...
*/

Fx.Slide = new Class({

	Extends: Fx,

	options: {
		mode: 'vertical',
		hideOverflow: true
	},

	initialize: function(element, options){
		this.addEvent('complete', function(){
			this.open = (this.wrapper['offset' + this.layout.capitalize()] != 0);
			if (this.open && Browser.Engine.webkit419) this.element.dispose().inject(this.wrapper);
		}, true);
		this.element = this.subject = document.id(element);
		this.parent(options);
		var wrapper = this.element.retrieve('wrapper');
		var styles = this.element.getStyles('margin', 'position', 'overflow');
		if (this.options.hideOverflow) styles = $extend(styles, {overflow: 'hidden'});
		this.wrapper = wrapper || new Element('div', {
			styles: styles
		}).wraps(this.element);
		this.element.store('wrapper', this.wrapper).setStyle('margin', 0);
		this.now = [];
		this.open = true;
	},

	vertical: function(){
		this.margin = 'margin-top';
		this.layout = 'height';
		this.offset = this.element.offsetHeight;
	},

	horizontal: function(){
		this.margin = 'margin-left';
		this.layout = 'width';
		this.offset = this.element.offsetWidth;
	},

	set: function(now){
		this.element.setStyle(this.margin, now[0]);
		this.wrapper.setStyle(this.layout, now[1]);
		return this;
	},

	compute: function(from, to, delta){
		return [0, 1].map(function(i){
			return Fx.compute(from[i], to[i], delta);
		});
	},

	start: function(how, mode){
		if (!this.check(how, mode)) return this;
		this[mode || this.options.mode]();
		var margin = this.element.getStyle(this.margin).toInt();
		var layout = this.wrapper.getStyle(this.layout).toInt();
		var caseIn = [[margin, layout], [0, this.offset]];
		var caseOut = [[margin, layout], [-this.offset, 0]];
		var start;
		switch (how){
			case 'in': start = caseIn; break;
			case 'out': start = caseOut; break;
			case 'toggle': start = (layout == 0) ? caseIn : caseOut;
		}
		return this.parent(start[0], start[1]);
	},

	slideIn: function(mode){
		return this.start('in', mode);
	},

	slideOut: function(mode){
		return this.start('out', mode);
	},

	hide: function(mode){
		this[mode || this.options.mode]();
		this.open = false;
		return this.set([-this.offset, 0]);
	},

	show: function(mode){
		this[mode || this.options.mode]();
		this.open = true;
		return this.set([0, this.offset]);
	},

	toggle: function(mode){
		return this.start('toggle', mode);
	}

});

Element.Properties.slide = {

	set: function(options){
		var slide = this.retrieve('slide');
		if (slide) slide.cancel();
		return this.eliminate('slide').store('slide:options', $extend({link: 'cancel'}, options));
	},

	get: function(options){
		if (options || !this.retrieve('slide')){
			if (options || !this.retrieve('slide:options')) this.set('slide', options);
			this.store('slide', new Fx.Slide(this, this.retrieve('slide:options')));
		}
		return this.retrieve('slide');
	}

};

Element.implement({

	slide: function(how, mode){
		how = how || 'toggle';
		var slide = this.get('slide'), toggle;
		switch (how){
			case 'hide': slide.hide(mode); break;
			case 'show': slide.show(mode); break;
			case 'toggle':
				var flag = this.retrieve('slide:flag', slide.open);
				slide[flag ? 'slideOut' : 'slideIn'](mode);
				this.store('slide:flag', !flag);
				toggle = true;
			break;
			default: slide.start(how, mode);
		}
		if (!toggle) this.eliminate('slide:flag');
		return this;
	}

});


/*
---

script: Fx.SmoothScroll.js

description: Class for creating a smooth scrolling effect to all internal links on the page.

license: MIT-style license

authors:
- Valerio Proietti

requires:
- core:1.2.4/Selectors
- /Fx.Scroll

provides: [Fx.SmoothScroll]

...
*/

var SmoothScroll = Fx.SmoothScroll = new Class({

	Extends: Fx.Scroll,

	initialize: function(options, context){
		context = context || document;
		this.doc = context.getDocument();
		var win = context.getWindow();
		this.parent(this.doc, options);
		this.links = $$(this.options.links || this.doc.links);
		var location = win.location.href.match(/^[^#]*/)[0] + '#';
		this.links.each(function(link){
			if (link.href.indexOf(location) != 0) {return;}
			var anchor = link.href.substr(location.length);
			if (anchor) this.useLink(link, anchor);
		}, this);
		if (!Browser.Engine.webkit419) {
			this.addEvent('complete', function(){
				win.location.hash = this.anchor;
			}, true);
		}
	},

	useLink: function(link, anchor){
		var el;
		link.addEvent('click', function(event){
			if (el !== false && !el) el = document.id(anchor) || this.doc.getElement('a[name=' + anchor + ']');
			if (el) {
				event.preventDefault();
				this.anchor = anchor;
				this.toElement(el).chain(function(){
					this.fireEvent('scrolledTo', [link, el]);
				}.bind(this));
				link.blur();
			}
		}.bind(this));
	}
});

/*
---

script: Drag.js

description: The base Drag Class. Can be used to drag and resize Elements using mouse events.

license: MIT-style license

authors:
- Valerio Proietti
- Tom Occhinno
- Jan Kassens

requires:
- core:1.2.4/Events
- core:1.2.4/Options
- core:1.2.4/Element.Event
- core:1.2.4/Element.Style
- /MooTools.More

provides: [Drag]

*/

var Drag = new Class({

	Implements: [Events, Options],

	options: {/*
		onBeforeStart: $empty(thisElement),
		onStart: $empty(thisElement, event),
		onSnap: $empty(thisElement)
		onDrag: $empty(thisElement, event),
		onCancel: $empty(thisElement),
		onComplete: $empty(thisElement, event),*/
		snap: 6,
		unit: 'px',
		grid: false,
		style: true,
		limit: false,
		handle: false,
		invert: false,
		preventDefault: false,
		stopPropagation: false,
		modifiers: {x: 'left', y: 'top'}
	},

	initialize: function(){
		var params = Array.link(arguments, {'options': Object.type, 'element': $defined});
		this.element = document.id(params.element);
		this.document = this.element.getDocument();
		this.setOptions(params.options || {});
		var htype = $type(this.options.handle);
		this.handles = ((htype == 'array' || htype == 'collection') ? $$(this.options.handle) : document.id(this.options.handle)) || this.element;
		this.mouse = {'now': {}, 'pos': {}};
		this.value = {'start': {}, 'now': {}};

		this.selection = (Browser.Engine.trident) ? 'selectstart' : 'mousedown';

		this.bound = {
			start: this.start.bind(this),
			check: this.check.bind(this),
			drag: this.drag.bind(this),
			stop: this.stop.bind(this),
			cancel: this.cancel.bind(this),
			eventStop: $lambda(false)
		};
		this.attach();
	},

	attach: function(){
		this.handles.addEvent('mousedown', this.bound.start);
		return this;
	},

	detach: function(){
		this.handles.removeEvent('mousedown', this.bound.start);
		return this;
	},

	start: function(event){
		if (event.rightClick) return;
		if (this.options.preventDefault) event.preventDefault();
		if (this.options.stopPropagation) event.stopPropagation();
		this.mouse.start = event.page;
		this.fireEvent('beforeStart', this.element);
		var limit = this.options.limit;
		this.limit = {x: [], y: []};
		for (var z in this.options.modifiers){
			if (!this.options.modifiers[z]) continue;
			if (this.options.style) this.value.now[z] = this.element.getStyle(this.options.modifiers[z]).toInt();
			else this.value.now[z] = this.element[this.options.modifiers[z]];
			if (this.options.invert) this.value.now[z] *= -1;
			this.mouse.pos[z] = event.page[z] - this.value.now[z];
			if (limit && limit[z]){
				for (var i = 2; i--; i){
					if ($chk(limit[z][i])) this.limit[z][i] = $lambda(limit[z][i])();
				}
			}
		}
		if ($type(this.options.grid) == 'number') this.options.grid = {x: this.options.grid, y: this.options.grid};
		this.document.addEvents({mousemove: this.bound.check, mouseup: this.bound.cancel});
		this.document.addEvent(this.selection, this.bound.eventStop);
	},

	check: function(event){
		if (this.options.preventDefault) event.preventDefault();
		var distance = Math.round(Math.sqrt(Math.pow(event.page.x - this.mouse.start.x, 2) + Math.pow(event.page.y - this.mouse.start.y, 2)));
		if (distance > this.options.snap){
			this.cancel();
			this.document.addEvents({
				mousemove: this.bound.drag,
				mouseup: this.bound.stop
			});
			this.fireEvent('start', [this.element, event]).fireEvent('snap', this.element);
		}
	},

	drag: function(event){
		if (this.options.preventDefault) event.preventDefault();
		this.mouse.now = event.page;
		for (var z in this.options.modifiers){
			if (!this.options.modifiers[z]) continue;
			this.value.now[z] = this.mouse.now[z] - this.mouse.pos[z];
			if (this.options.invert) this.value.now[z] *= -1;
			if (this.options.limit && this.limit[z]){
				if ($chk(this.limit[z][1]) && (this.value.now[z] > this.limit[z][1])){
					this.value.now[z] = this.limit[z][1];
				} else if ($chk(this.limit[z][0]) && (this.value.now[z] < this.limit[z][0])){
					this.value.now[z] = this.limit[z][0];
				}
			}
			if (this.options.grid[z]) this.value.now[z] -= ((this.value.now[z] - (this.limit[z][0]||0)) % this.options.grid[z]);
			if (this.options.style) {
				this.element.setStyle(this.options.modifiers[z], this.value.now[z] + this.options.unit);
			} else {
				this.element[this.options.modifiers[z]] = this.value.now[z];
			}
		}
		this.fireEvent('drag', [this.element, event]);
	},

	cancel: function(event){
		this.document.removeEvent('mousemove', this.bound.check);
		this.document.removeEvent('mouseup', this.bound.cancel);
		if (event){
			this.document.removeEvent(this.selection, this.bound.eventStop);
			this.fireEvent('cancel', this.element);
		}
	},

	stop: function(event){
		this.document.removeEvent(this.selection, this.bound.eventStop);
		this.document.removeEvent('mousemove', this.bound.drag);
		this.document.removeEvent('mouseup', this.bound.stop);
		if (event) this.fireEvent('complete', [this.element, event]);
	}

});

Element.implement({

	makeResizable: function(options){
		var drag = new Drag(this, $merge({modifiers: {x: 'width', y: 'height'}}, options));
		this.store('resizer', drag);
		return drag.addEvent('drag', function(){
			this.fireEvent('resize', drag);
		}.bind(this));
	}

});


/*
---

script: Drag.Move.js

description: A Drag extension that provides support for the constraining of draggables to containers and droppables.

license: MIT-style license

authors:
- Valerio Proietti
- Tom Occhinno
- Jan Kassens
- Aaron Newton
- Scott Kyle

requires:
- core:1.2.4/Element.Dimensions
- /Drag

provides: [Drag.Move]

...
*/

Drag.Move = new Class({

	Extends: Drag,

	options: {/*
		onEnter: $empty(thisElement, overed),
		onLeave: $empty(thisElement, overed),
		onDrop: $empty(thisElement, overed, event),*/
		droppables: [],
		container: false,
		precalculate: false,
		includeMargins: true,
		checkDroppables: true
	},

	initialize: function(element, options){
		this.parent(element, options);
		element = this.element;
		
		this.droppables = $$(this.options.droppables);
		this.container = document.id(this.options.container);
		
		if (this.container && $type(this.container) != 'element')
			this.container = document.id(this.container.getDocument().body);
		
		var styles = element.getStyles('left', 'right', 'position');
		if (styles.left == 'auto' || styles.top == 'auto')
			element.setPosition(element.getPosition(element.getOffsetParent()));
		
		if (styles.position == 'static')
			element.setStyle('position', 'absolute');

		this.addEvent('start', this.checkDroppables, true);

		this.overed = null;
	},

	start: function(event){
		if (this.container) this.options.limit = this.calculateLimit();
		
		if (this.options.precalculate){
			this.positions = this.droppables.map(function(el){
				return el.getCoordinates();
			});
		}
		
		this.parent(event);
	},
	
	calculateLimit: function(){
		var offsetParent = this.element.getOffsetParent(),
			containerCoordinates = this.container.getCoordinates(offsetParent),
			containerBorder = {},
			elementMargin = {},
			elementBorder = {},
			containerMargin = {},
			offsetParentPadding = {};

		['top', 'right', 'bottom', 'left'].each(function(pad){
			containerBorder[pad] = this.container.getStyle('border-' + pad).toInt();
			elementBorder[pad] = this.element.getStyle('border-' + pad).toInt();
			elementMargin[pad] = this.element.getStyle('margin-' + pad).toInt();
			containerMargin[pad] = this.container.getStyle('margin-' + pad).toInt();
			offsetParentPadding[pad] = offsetParent.getStyle('padding-' + pad).toInt();
		}, this);

		var width = this.element.offsetWidth + elementMargin.left + elementMargin.right,
			height = this.element.offsetHeight + elementMargin.top + elementMargin.bottom,
			left = 0,
			top = 0,
			right = containerCoordinates.right - containerBorder.right - width,
			bottom = containerCoordinates.bottom - containerBorder.bottom - height;

		if (this.options.includeMargins){
			left += elementMargin.left;
			top += elementMargin.top;
		} else {
			right += elementMargin.right;
			bottom += elementMargin.bottom;
		}
		
		if (this.element.getStyle('position') == 'relative'){
			var coords = this.element.getCoordinates(offsetParent);
			coords.left -= this.element.getStyle('left').toInt();
			coords.top -= this.element.getStyle('top').toInt();
			
			left += containerBorder.left - coords.left;
			top += containerBorder.top - coords.top;
			right += elementMargin.left - coords.left;
			bottom += elementMargin.top - coords.top;
			
			if (this.container != offsetParent){
				left += containerMargin.left + offsetParentPadding.left;
				top += (Browser.Engine.trident4 ? 0 : containerMargin.top) + offsetParentPadding.top;
			}
		} else {
			left -= elementMargin.left;
			top -= elementMargin.top;
			
			if (this.container == offsetParent){
				right -= containerBorder.left;
				bottom -= containerBorder.top;
			} else {
				left += containerCoordinates.left + containerBorder.left;
				top += containerCoordinates.top + containerBorder.top;
			}
		}
		
		return {
			x: [left, right],
			y: [top, bottom]
		};
	},

	checkAgainst: function(el, i){
		el = (this.positions) ? this.positions[i] : el.getCoordinates();
		var now = this.mouse.now;
		return (now.x > el.left && now.x < el.right && now.y < el.bottom && now.y > el.top);
	},

	checkDroppables: function(){
		var overed = this.droppables.filter(this.checkAgainst, this).getLast();
		if (this.overed != overed){
			if (this.overed) this.fireEvent('leave', [this.element, this.overed]);
			if (overed) this.fireEvent('enter', [this.element, overed]);
			this.overed = overed;
		}
	},

	drag: function(event){
		this.parent(event);
		if (this.options.checkDroppables && this.droppables.length) this.checkDroppables();
	},

	stop: function(event){
		this.checkDroppables();
		this.fireEvent('drop', [this.element, this.overed, event]);
		this.overed = null;
		return this.parent(event);
	}

});

Element.implement({

	makeDraggable: function(options){
		var drag = new Drag.Move(this, options);
		this.store('dragger', drag);
		return drag;
	}

});


/*
---

script: Slider.js

description: Class for creating horizontal and vertical slider controls.

license: MIT-style license

authors:
- Valerio Proietti

requires:
- core:1.2.4/Element.Dimensions
- /Class.Binds
- /Drag
- /Element.Dimensions
- /Element.Measure

provides: [Slider]

...
*/

var Slider = new Class({

	Implements: [Events, Options],

	Binds: ['clickedElement', 'draggedKnob', 'scrolledElement'],

	options: {/*
		onTick: $empty(intPosition),
		onChange: $empty(intStep),
		onComplete: $empty(strStep),*/
		onTick: function(position){
			if (this.options.snap) position = this.toPosition(this.step);
			this.knob.setStyle(this.property, position);
		},
		initialStep: 0,
		snap: false,
		offset: 0,
		range: false,
		wheel: false,
		steps: 100,
		mode: 'horizontal'
	},

	initialize: function(element, knob, options){
		this.setOptions(options);
		this.element = document.id(element);
		this.knob = document.id(knob);
		this.previousChange = this.previousEnd = this.step = -1;
		var offset, limit = {}, modifiers = {'x': false, 'y': false};
		switch (this.options.mode){
			case 'vertical':
				this.axis = 'y';
				this.property = 'top';
				offset = 'offsetHeight';
				break;
			case 'horizontal':
				this.axis = 'x';
				this.property = 'left';
				offset = 'offsetWidth';
		}
		
		this.full = this.element.measure(function(){ 
			this.half = this.knob[offset] / 2; 
			return this.element[offset] - this.knob[offset] + (this.options.offset * 2); 
		}.bind(this));
		
		this.min = $chk(this.options.range[0]) ? this.options.range[0] : 0;
		this.max = $chk(this.options.range[1]) ? this.options.range[1] : this.options.steps;
		this.range = this.max - this.min;
		this.steps = this.options.steps || this.full;
		this.stepSize = Math.abs(this.range) / this.steps;
		this.stepWidth = this.stepSize * this.full / Math.abs(this.range) ;

		this.knob.setStyle('position', 'relative').setStyle(this.property, this.options.initialStep ? this.toPosition(this.options.initialStep) : - this.options.offset);
		modifiers[this.axis] = this.property;
		limit[this.axis] = [- this.options.offset, this.full - this.options.offset];

		var dragOptions = {
			snap: 0,
			limit: limit,
			modifiers: modifiers,
			onDrag: this.draggedKnob,
			onStart: this.draggedKnob,
			onBeforeStart: (function(){
				this.isDragging = true;
			}).bind(this),
			onCancel: function() {
				this.isDragging = false;
			}.bind(this),
			onComplete: function(){
				this.isDragging = false;
				this.draggedKnob();
				this.end();
			}.bind(this)
		};
		if (this.options.snap){
			dragOptions.grid = Math.ceil(this.stepWidth);
			dragOptions.limit[this.axis][1] = this.full;
		}

		this.drag = new Drag(this.knob, dragOptions);
		this.attach();
	},

	attach: function(){
		this.element.addEvent('mousedown', this.clickedElement);
		if (this.options.wheel) this.element.addEvent('mousewheel', this.scrolledElement);
		this.drag.attach();
		return this;
	},

	detach: function(){
		this.element.removeEvent('mousedown', this.clickedElement);
		this.element.removeEvent('mousewheel', this.scrolledElement);
		this.drag.detach();
		return this;
	},

	set: function(step){
		if (!((this.range > 0) ^ (step < this.min))) step = this.min;
		if (!((this.range > 0) ^ (step > this.max))) step = this.max;

		this.step = Math.round(step);
		this.checkStep();
		this.fireEvent('tick', this.toPosition(this.step));
		this.end();
		return this;
	},

	clickedElement: function(event){
		if (this.isDragging || event.target == this.knob) return;

		var dir = this.range < 0 ? -1 : 1;
		var position = event.page[this.axis] - this.element.getPosition()[this.axis] - this.half;
		position = position.limit(-this.options.offset, this.full -this.options.offset);

		this.step = Math.round(this.min + dir * this.toStep(position));
		this.checkStep();
		this.fireEvent('tick', position);
		this.end();
	},

	scrolledElement: function(event){
		var mode = (this.options.mode == 'horizontal') ? (event.wheel < 0) : (event.wheel > 0);
		this.set(mode ? this.step - this.stepSize : this.step + this.stepSize);
		event.stop();
	},

	draggedKnob: function(){
		var dir = this.range < 0 ? -1 : 1;
		var position = this.drag.value.now[this.axis];
		position = position.limit(-this.options.offset, this.full -this.options.offset);
		this.step = Math.round(this.min + dir * this.toStep(position));
		this.checkStep();
	},

	checkStep: function(){
		if (this.previousChange != this.step){
			this.previousChange = this.step;
			this.fireEvent('change', this.step);
		}
	},

	end: function(){
		if (this.previousEnd !== this.step){
			this.previousEnd = this.step;
			this.fireEvent('complete', this.step + '');
		}
	},

	toStep: function(position){
		var step = (position + this.options.offset) * this.stepSize / this.full * this.steps;
		return this.options.steps ? Math.round(step -= step % this.stepSize) : step;
	},

	toPosition: function(step){
		return (this.full * Math.abs(this.min - step)) / (this.steps * this.stepSize) - this.options.offset;
	}

});

/*
---

script: Sortables.js

description: Class for creating a drag and drop sorting interface for lists of items.

license: MIT-style license

authors:
- Tom Occhino

requires:
- /Drag.Move

provides: [Slider]

...
*/

var Sortables = new Class({

	Implements: [Events, Options],

	options: {/*
		onSort: $empty(element, clone),
		onStart: $empty(element, clone),
		onComplete: $empty(element),*/
		snap: 4,
		opacity: 1,
		clone: false,
		revert: false,
		handle: false,
		constrain: false
	},

	initialize: function(lists, options){
		this.setOptions(options);
		this.elements = [];
		this.lists = [];
		this.idle = true;

		this.addLists($$(document.id(lists) || lists));
		if (!this.options.clone) this.options.revert = false;
		if (this.options.revert) this.effect = new Fx.Morph(null, $merge({duration: 250, link: 'cancel'}, this.options.revert));
	},

	attach: function(){
		this.addLists(this.lists);
		return this;
	},

	detach: function(){
		this.lists = this.removeLists(this.lists);
		return this;
	},

	addItems: function(){
		Array.flatten(arguments).each(function(element){
			this.elements.push(element);
			var start = element.retrieve('sortables:start', this.start.bindWithEvent(this, element));
			(this.options.handle ? element.getElement(this.options.handle) || element : element).addEvent('mousedown', start);
		}, this);
		return this;
	},

	addLists: function(){
		Array.flatten(arguments).each(function(list){
			this.lists.push(list);
			this.addItems(list.getChildren());
		}, this);
		return this;
	},

	removeItems: function(){
		return $$(Array.flatten(arguments).map(function(element){
			this.elements.erase(element);
			var start = element.retrieve('sortables:start');
			(this.options.handle ? element.getElement(this.options.handle) || element : element).removeEvent('mousedown', start);
			
			return element;
		}, this));
	},

	removeLists: function(){
		return $$(Array.flatten(arguments).map(function(list){
			this.lists.erase(list);
			this.removeItems(list.getChildren());
			
			return list;
		}, this));
	},

	getClone: function(event, element){
		if (!this.options.clone) return new Element('div').inject(document.body);
		if ($type(this.options.clone) == 'function') return this.options.clone.call(this, event, element, this.list);
		return element.clone(true).setStyles({
			margin: '0px',
			position: 'absolute',
			visibility: 'hidden',
			'width': element.getStyle('width')
		}).inject(this.list).setPosition(element.getPosition(element.getOffsetParent()));
	},

	getDroppables: function(){
		var droppables = this.list.getChildren();
		if (!this.options.constrain) droppables = this.lists.concat(droppables).erase(this.list);
		return droppables.erase(this.clone).erase(this.element);
	},

	insert: function(dragging, element){
		var where = 'inside';
		if (this.lists.contains(element)){
			this.list = element;
			this.drag.droppables = this.getDroppables();
		} else {
			where = this.element.getAllPrevious().contains(element) ? 'before' : 'after';
		}
		this.element.inject(element, where);
		this.fireEvent('sort', [this.element, this.clone]);
	},

	start: function(event, element){
		if (!this.idle) return;
		this.idle = false;
		this.element = element;
		this.opacity = element.get('opacity');
		this.list = element.getParent();
		this.clone = this.getClone(event, element);

		this.drag = new Drag.Move(this.clone, {
			snap: this.options.snap,
			container: this.options.constrain && this.element.getParent(),
			droppables: this.getDroppables(),
			onSnap: function(){
				event.stop();
				this.clone.setStyle('visibility', 'visible');
				this.element.set('opacity', this.options.opacity || 0);
				this.fireEvent('start', [this.element, this.clone]);
			}.bind(this),
			onEnter: this.insert.bind(this),
			onCancel: this.reset.bind(this),
			onComplete: this.end.bind(this)
		});

		this.clone.inject(this.element, 'before');
		this.drag.start(event);
	},

	end: function(){
		this.drag.detach();
		this.element.set('opacity', this.opacity);
		if (this.effect){
			var dim = this.element.getStyles('width', 'height');
			var pos = this.clone.computePosition(this.element.getPosition(this.clone.offsetParent));
			this.effect.element = this.clone;
			this.effect.start({
				top: pos.top,
				left: pos.left,
				width: dim.width,
				height: dim.height,
				opacity: 0.25
			}).chain(this.reset.bind(this));
		} else {
			this.reset();
		}
	},

	reset: function(){
		this.idle = true;
		this.clone.destroy();
		this.fireEvent('complete', this.element);
	},

	serialize: function(){
		var params = Array.link(arguments, {modifier: Function.type, index: $defined});
		var serial = this.lists.map(function(list){
			return list.getChildren().map(params.modifier || function(element){
				return element.get('id');
			}, this);
		}, this);

		var index = params.index;
		if (this.lists.length == 1) index = 0;
		return $chk(index) && index >= 0 && index < this.lists.length ? serial[index] : serial;
	}

});
