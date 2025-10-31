let headers = [];
let rows = [];

self.onmessage = (event) => {
    const { type } = event.data;
    if (type === 'initData') {
        headers = event.data.headers || [];
        rows = event.data.rows || [];
        self.postMessage({ type: 'ready', total: rows.length });
    } else if (type === 'search') {
        try {
            const { query, options } = event.data;
            const start = performance.now();
            const matches = executeSearch(query, options || {});
            const duration = performance.now() - start;
            self.postMessage({ type: 'searchResults', matches, duration });
        } catch (error) {
            self.postMessage({ type: 'error', message: error.message });
        }
    } else if (type === 'reset') {
        headers = [];
        rows = [];
    }
};

function executeSearch(query, options) {
    if (!rows.length) {
        return [];
    }

    const normalized = (query || '').trim();
    if (!normalized) {
        return rows.map((_, index) => index);
    }

    const expression = buildExpression(normalized);
    const postfix = infixToPostfix(expression);
    if (!postfix.length) {
        return rows.map((_, index) => index);
    }

    const matches = [];
    for (let index = 0; index < rows.length; index += 1) {
        const rowValues = rows[index];
        if (evaluatePostfix(postfix, rowValues, options)) {
            matches.push(index);
        }
    }
    return matches;
}

function buildExpression(query) {
    const replaced = query
        .replace(/[\u201C\u201D]/g, '"')
        .replace(/[,;]/g, ' OR ')
        .replace(/\s+OR\s+/gi, ' OR ')
        .replace(/\s+AND\s+/gi, ' AND ')
        .replace(/\s+NOT\s+/gi, ' NOT ');

    const tokens = [];
    const rawTokens = replaced.match(/"[^"]+"|\(|\)|[^\s]+/g) || [];

    rawTokens.forEach((token) => {
        if (!token.trim()) {
            return;
        }
        const upper = token.toUpperCase();
        if (upper === 'AND' || upper === 'OR' || upper === 'NOT') {
            tokens.push({ type: 'op', value: upper });
            return;
        }
        if (token === '(' || token === ')') {
            tokens.push({ type: 'paren', value: token });
            return;
        }
        if (token.startsWith('-') && token.length > 1) {
            tokens.push({ type: 'op', value: 'NOT' });
            tokens.push({ type: 'term', value: stripQuotes(token.slice(1)) });
            return;
        }
        tokens.push({ type: 'term', value: stripQuotes(token) });
    });

    const enriched = [];
    for (let i = 0; i < tokens.length; i += 1) {
        const current = tokens[i];
        const previous = tokens[i - 1];
        if (current.type === 'term' && previous && (previous.type === 'term' || previous.value === ')')) {
            enriched.push({ type: 'op', value: 'AND' });
        } else if (current.value === '(' && previous && (previous.type === 'term' || previous.value === ')')) {
            enriched.push({ type: 'op', value: 'AND' });
        }
        enriched.push(current);
    }

    return enriched;
}

function stripQuotes(value) {
    if (!value) {
        return '';
    }
    return value.replace(/^"|"$/g, '');
}

function infixToPostfix(tokens) {
    const output = [];
    const operators = [];
    const precedence = { NOT: 3, AND: 2, OR: 1 };

    tokens.forEach((token) => {
        if (token.type === 'term') {
            output.push(token);
            return;
        }
        if (token.type === 'op') {
            while (operators.length) {
                const last = operators[operators.length - 1];
                if (last.type === 'op' && precedence[last.value] >= precedence[token.value]) {
                    output.push(operators.pop());
                } else {
                    break;
                }
            }
            operators.push(token);
            return;
        }
        if (token.value === '(') {
            operators.push(token);
            return;
        }
        if (token.value === ')') {
            while (operators.length && operators[operators.length - 1].value !== '(') {
                output.push(operators.pop());
            }
            operators.pop();
        }
    });

    while (operators.length) {
        output.push(operators.pop());
    }
    return output;
}

function evaluatePostfix(postfix, rowValues, options) {
    const stack = [];
    postfix.forEach((token) => {
        if (token.type === 'term') {
            const result = matchRow(rowValues, token.value, options);
            stack.push(result);
            return;
        }
        if (token.type === 'op') {
            if (token.value === 'NOT') {
                const value = stack.pop() ?? false;
                stack.push(!value);
            } else if (token.value === 'AND') {
                const right = stack.pop() ?? false;
                const left = stack.pop() ?? false;
                stack.push(left && right);
            } else if (token.value === 'OR') {
                const right = stack.pop() ?? false;
                const left = stack.pop() ?? false;
                stack.push(left || right);
            }
        }
    });
    return stack.pop() ?? false;
}

function matchRow(rowValues, term, options) {
    if (!term) {
        return true;
    }
    const { caseSensitive, exactMatch } = options;
    const target = caseSensitive ? term : term.toLowerCase();
    for (let i = 0; i < rowValues.length; i += 1) {
        const rawValue = rowValues[i];
        if (rawValue === undefined || rawValue === null) {
            continue;
        }
        const cellValue = String(rawValue);
        const comparable = caseSensitive ? cellValue : cellValue.toLowerCase();
        if (exactMatch) {
            if (comparable === target) {
                return true;
            }
        } else if (comparable.includes(target)) {
            return true;
        }
    }
    return false;
}
