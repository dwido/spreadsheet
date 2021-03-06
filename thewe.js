we = {}

$not = function(f) {
	return function(x) {
		return !f(x)
	}
}

String.implement({
	beginsWith: function(pre) {
		return this.substring(0, pre.length) == pre
	}
})

$begins = function(prefixes) {
	prefixes = $splat(prefixes)

	return function(str) {
		return prefixes.some(function(prefix) {
			return str.beginsWith(prefix)
		})
	}
}

Array.implement({
	getAllButLast: function() {
		return this.filter(function(value, i) {
			return i < this.length - 1
		}.bind(this))
	}
})

we.delta = {}

we.submitChanges = function() {
	wave.getState().submitDelta(we.delta)
	we.delta = {}
}

we.State = new Class({
	Extends: Hash,

	initialize: function(cursorPath) {
		this._cursorPath = cursorPath || ''
	},

	set: function(key, value, submit) {
		if (['object', 'hash'].contains($type(value)))
			we.flattenState(value, this._cursorPath + (key ? (key + '.') : ''), we.delta)
		else
			we.delta[this._cursorPath + key] = value

		if (submit)
			we.submitChanges()

		return this
	},

	unset: function(key, submit) {
		var oldValue = this[key]

		if (['object', 'hash'].contains($type(oldValue)))
			oldValue.getKeys().each(function(subkey) {
				oldValue.unset(subkey)
			})
		else
			we.delta[this._cursorPath + key] = null

		if (submit)
			we.submitChanges()

		return this
	},		

	getKeys: function() {
		return this.parent().filter($not($begins(['_', '$', 'caller' /* $todo */])))
	}
})

Hash.implement({
	filterKeys: function(filter) {
		return this.filter(function(value, key) {
			return filter(key)
		})
	}
})

we.deepenState = function(state) {
	var result = new we.State()
	$H(state.state_).filterKeys($not($begins('$'))).each(function(value, key) {
		var cursor = result
		var tokens = key.split('.')
		var cursorPath = ''

		tokens.getAllButLast().each(function(token) {
			cursorPath += token + '.';

			if (!cursor[token])
				cursor[token] = new we.State(cursorPath)

			cursor = cursor[token]
		})

		cursor[tokens.getLast()] = value
	})

	return result
}

we.flattenState = function(state, cursorPath, into) {
	cursorPath = cursorPath || ''
	into = into || $H()

	$H(state).each(function(value, key) {
		if (['object', 'hash'].contains($type(value)))
			we.flattenState(value, cursorPath + key + '.', into)
		else
			into[cursorPath + key] = value
	})

	return into
}

we.computeState = function() {
	return we.state = we.deepenState(wave.getState())
}

function main() {
	if (wave && wave.isInWaveContainer()) {
                window.addEvent('keypress', function(event) {
			if (event.alt && event.control) {
				var key = String.fromCharCode(event.event.charCode)
				if (key == 'e') {
					// $fix?
					we.state.set(null, 
						JSON.parse(prompt("Gadget state", 
							JSON.stringify(we.computeState()))))
				}
			}
		})

		wave.setStateCallback(function() {
			if (!we.isEvaled) {
				eval(wave.getState().get('_view'))
				we.isEvaled = true
			}

			stateUpdated(we.computeState())
		})
	}
}

gadgets.util.registerOnLoadHandler(main)


