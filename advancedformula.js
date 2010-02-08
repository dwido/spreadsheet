var AdvancedFormula = new Class({
    initialize: function(owner){
        this.ownerGrid = owner;
    },
    
    calculate: function(formula){
        var name = this.getName(formula);
        var arguments = this.getArguments(formula);
        var parsedArguments = this.parseArguments(arguments);// arguments can consist of range and cells parse them to one array of all the cells that contained
        var arrValues = this.argumentsToValues(parsedArguments);
        switch(name.toUpperCase())
        {
            case 'SUM':
                    return this.sum(arrValues);
            case 'AVERAGE':
                    return this.average(arrValues);
            case 'PRODUCT':
                    return this.product(arrValues);
            case 'MIN':
                    return this.min(arrValues);
            case 'MAX':
                    return this.max(arrValues);
        }
    },

    getName : function(formula){
        return formula.substr(this.startsWEqual(formula)?1:0, formula.indexOf('(') - 1);
    },
    
    getArguments : function(formula){
        var firstIndex = formula.indexOf('(');
        var lastIndex = formula.lastIndexOf(')');
        return formula.substr(firstIndex + 1, lastIndex - firstIndex - 1);
    },
    
    startsWEqual: function(text){
        return text.substr(0,1) == '=';
    },
    
    isAdvancedFormula: function(text){
        
        var pLoc = text.indexOf('(');
        if(pLoc!=-1 && this.startsWEqual(text))
            if(this.ownerGrid.advancedFormula.contains(text.toUpperCase().substr(1,pLoc-1),','))
                return true;

        return false;
    },
    
    sum: function(arrValues){
        //var arrCells = arguments.split(',');
        var sum=0;
        arrValues.each(function(item){
            sum += parseInt(item);
        })
        return sum;
    },
    
    average: function(arrValues){
        //var arrCells = arguments.split(',');
        var sum=0;
        arrValues.each(function(item){
            sum += parseInt(item);
        })
        return sum/arrValues.length;
    },
    
    product: function(arrValues){
        //var arrCells = arguments.split(',');
        var retval=1;
        arrValues.each(function(item){
            retval *= parseInt(item);
        })
        return retval;
    },
    
    max: function(arrValues){
        var retval = arrValues[0];
        arrValues.each(function(item){
            if (retval < parseInt(item)) retval = parseInt(item);
        })
        return retval;
    },
    
    min: function(arrValues){
        var retval = arrValues[0];
        arrValues.each(function(item){
            if (retval > parseInt(item)) retval = parseInt(item);
        })
        return retval;
    },
    
    // "(A7:B8,G5,A1:B2)" --> ['A7','A8','B7','B8','G5','A1','A2','B1','B2']
    parseArguments: function(arguments){
        var arrCells = arguments.split(',');
        var arrParseArguments = [];
        arrCells.each(function(item){
            if(this.ownerGrid.isRange(item))
                arrParseArguments.combine(this.ownerGrid.parseRange(item));
            else
                arrParseArguments[arrParseArguments.length] = item;
        }.bind(this))
        
        return arrParseArguments;
    },
    
    argumentsToValues: function(parsedArguments){
        var arr = [];
        parsedArguments.each(function(a){
            arr[arr.length] = this.ownerGrid.getElementByCaption(a).get('text');
        }.bind(this))
        return arr;
    }
}) // end of class