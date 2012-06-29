var exec = require('child_process').exec;
var sax = require('sax');

var objectIndex = 0;

function queryPS(comType, query, callback) {
	var command = '[console]::OutputEncoding = New-Object -typename System.Text.UTF8Encoding; ((new-object -com '+comType+').'+query+' | ConvertTo-XML).InnerXml';
	//command = new Buffer(command);
	//command = command.toString('base64');
	//exec('PowerShell.exe -NoProfile -NonInteractive -NoLogo -EncodedCommand "'+command+'"', {encoding: 'utf8'}, function(error, stdout, stderr) {
	exec('PowerShell.exe -NoProfile -NonInteractive -NoLogo -Command "'+command+'"', {encoding: 'utf8'}, function(error, stdout, stderr) {
		if(error) {
			return callback(error);
		} else if(stderr) {
			return callback(new Error(stderr));
		}
		callback(null, stdout.replace(/^[^<]+/, '').replace(/(\r\n)+/g, ''));
	}).stdin.end();
}

function COMObject(session, path, parent, type, properties, subObjects) {
	var _this = this;
	this.session = session;
	if(path) {
		this.path = path;
	}
	if(parent) {
		this.parent = parent;
	}
	if(type) {
		this.type = type;
	}
	if(properties) {
		this.properties = properties;
	}
	if(subObjects) {
		this.query = this.query.bind(this);
		subObjects.forEach(function(name) {
			_this.query[name] = function(callback) {
				return _this.query(name, callback);
			};
		});
	}
}

COMObject.prototype.fullPath = function() {
	if(this.parent) {
		return this.parent.fullPath() + this.path + '.';
	}
	if(this.path) {
		return this.path + '.';
	}
	return '';
};

COMObject.prototype.parseResult = function(text, query, callback) {
	var _this = this;
	var parser = sax.parser(true);
	var currentNode = null;
	var properties = {};
	var subObjects = [];
	var error = null;
	var type = null;
	parser.onopentag = function(node) {
		node.text = '';
		node.parentNode = currentNode;
		currentNode = node;
	};
	parser.onclosetag = function() {
		if(currentNode.name === 'Object') {
			type = currentNode.attributes.Type;
		} else if(currentNode.name === 'Property') {
			if(!currentNode.attributes.Type) {
				error = new Error('Property '+currentNode.attributes.Name+' does not have a type');
				return;
			}
			var propertyType = currentNode.attributes.Type.replace(/^System\./, '');
			var name = currentNode.attributes.Name;
			var value = currentNode.text;

			if(propertyType === '__ComObject' || propertyType === 'Object') {
				subObjects.push(name);
			} else {
				if(propertyType === 'String') {
					//String
					value = value;
				} else if(propertyType === 'Decimal') {
					value = parseFloat(value);
				} else if(propertyType === 'Boolean') {
					value = value === 'True';
				} else if(propertyType === 'DateTime') {
					value = new Date(value+'Z');
				} else if(propertyType === 'Nil') {
					value = null;
				} else if(propertyType.indexOf('Int') === 0) {
					value = parseInt(value, 10);
				} else {
					error = new Error('Unknown attribute type '+propertyType);
				}
				properties[name] = value;
			}
		}
		currentNode = currentNode.parentNode;
	};
	parser.ontext = function(text) {
		if(currentNode) {
			currentNode.text += text;
		}
	};
	parser.onerror = function(e) {
		error = e;
	};
	parser.onend = function() {
		var result = error ? null : new COMObject(_this.session, query, _this, type, properties, subObjects);
		callback(error, result);
	};
	parser.write(text).end();
};


COMObject.prototype.query = function(query, callback) {
	var _this = this;
	this.queryText(query, function(error, text) {
		if(error) {
			return callback(error);
		}
		//Parse object
		_this.parseResult(text, query, callback);
	});
};

COMObject.prototype.queryMethod = function(methodName) {
	var args = [].slice.call(arguments, 1, -1);
	var i = 0;

	for(i=0;i<args.length;i++) {
		var value = args[i];
		if(value === null || value === undefined) {
			value = 'nil';
		} else if(value.constructor === String) {
			value = "'"+value.replace(/'/g, '"')+"'";
		} else if(value.constructor === Boolean) {
			value = value ? '$True' : '$False';
		} else {
			value = ''+value;
		}
		args[i] = value;
	}
	args = args.join(', ');
	var callback = arguments[arguments.length-1];
	methodName+='('+args+')';
	this.query(methodName, callback);
};

COMObject.prototype.queryText = function(query, callback) {
	queryPS(this.session.comType, this.fullPath()+query, callback);
};

function COMSession(comType) {
	var _this = this;
	this.comType = comType;
	COMObject.call(this, this);
	this.callbackStack = [];
}

COMSession.prototype.__proto__ = COMObject.prototype;

exports.COMSession = COMSession;