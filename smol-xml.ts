

export class SmolXMLNode {
  constructor(public tag: string) { }

  attrs: Record<string, string> = {}
  children: (SmolXMLNode | string)[] = []

  private _textContent(res: string[]) {
    for (const child of this.children) {
      if (child instanceof SmolXMLNode) {
        child._textContent(res)
      } else {
        res.push(child)
      }
    }
  }

  get textContent(): string {
    const res: string[] = []
    this._textContent(res)
    return res.join("")
  }
}

let spaces = new Set([
  " ",
  "\t",
  "\n",
  "\r",
])

const re_xml_chars = /&(lt|gt|quot|a(?=mp|pos)|#\d+);/g

export function decodeXMLChars(str: string): string {
  return str.replace(re_xml_chars, (_, m) => {
    switch (m[0]) {
      case "l": return "<"
      case "g": return ">"
      case "q": return '"'
      case "#": return String.fromCharCode(Number(m.slice(1)))
      case "a": return m[1] === "m" ? "&" : "'"
      default: return m
    }
  })
}

export function streamXML(_target: string | string[], file: string, on_node: (node: SmolXMLNode) => void) {
  // Pre-allocate the stack

  let stack: SmolXMLNode[] = new Array(100)
  let stack_index = -1
  const single_target = !Array.isArray(_target)
  let target = Array.isArray(_target) ? new Set(_target) : _target

  // Pre-allocate the current node
  let current_node: SmolXMLNode | null = null as SmolXMLNode | null
  current_node = null
  let last_text = 0

  let i = 0
  let len = file.length

  while (i < len) {

    last_text = i
    let needs_decode = false

    while (i < len && file[i] !== '<') {
      if (file[i] === '&') {
        needs_decode = true
      }
      i++
    }

    // We're entering node context, so push the text to the current node
    if (current_node != null && last_text < i) {
      const text = file.slice(last_text, i)
      current_node.children.push(needs_decode ? decodeXMLChars(text) : text)
    }

    // Stop if we've reached the end of the chunk
    if (i >= len) {
      break
    }

    if (file[i] === '<') {

      // This is actually a close tag, the current node will be closed
      if (i + 1 < len && file[i + 1] === '/') {
        i++
        // advance to the closing >
        while (i < len && file[i] !== '>') {
          i++
        }

        if (stack_index === 0) {
          on_node(current_node!)
          stack_index = -1
          current_node = null
        } else if (stack_index > 0) {
          stack[stack_index] = null as unknown as SmolXMLNode
          stack_index--
          current_node = stack[stack_index]
        }

        i++ // Advance to the next character after the closing >
        continue
      }

      // We're going towards the open tag state
      // advance to the next character
      i++
      let tag_start = i
      //
      while (i < len && !spaces.has(file[i]) && file[i] !== '>' && file[i] !== '/') {
        i++
      }

      let _tag_name = file.slice(tag_start, i)

      // if (i < len && file[i] === '>') {
      //   i++ // Advance to the next character after the tag name
      // }

      if (current_node == null && (single_target ? _tag_name === target : (target as Set<string>).has(_tag_name))) {
        // We're matchin a root node we're looking for
        current_node = new SmolXMLNode(_tag_name)
        stack_index++
        stack[stack_index] = current_node
      } else if (current_node != null) {
        // push a new node
        const new_node = new SmolXMLNode(_tag_name)
        current_node.children.push(new_node)
        current_node = new_node
        stack_index++
        stack[stack_index] = new_node
        current_node = new_node
      }

      // Advance to the next non-space character
      while (i < len && spaces.has(file[i])) {
        i++
      }

      // Now, parse the attributes one by one
      while (i < len && file[i] !== '>' && file[i] !== '/') {

        let attr_start = i
        // An attribute name can stop at =, >, or a space, because the attribute may just be an empty attribute
        while (i < len && file[i] !== '=' && !spaces.has(file[i]) && file[i] !== '>') {
          i++
        }
        let attr_name = file.slice(attr_start, i)

        if (i >= len) {
          continue
        }

        let attr_value = ""
        // This will be a value attribute
        if (file[i] === '=') {
          i++
          let delimiter = file[i]
          i++
          let attr_value_start = i
          let needs_decode = false
          while (i < len && file[i] !== delimiter) {
            if (file[i] === "\\" && file[i + 1] === delimiter) {
              i++
            }
            if (file[i] === '&') {
              needs_decode = true
            }
            i++
          }
          attr_value = file.slice(attr_value_start, i)
          if (needs_decode) {
            attr_value = decodeXMLChars(attr_value)
          }
          i++ // Advance to the next character after the delimiter
        }

        if (current_node != null) {
          current_node.attrs[attr_name] = attr_value
        }

        // Advance to the next non-space character
        while (i < len && spaces.has(file[i])) {
          i++
        }
      }

      if (i < len && file[i] === '/') {
        i++
        if (current_node != null) {
          // This pops the node as it is self closing
          stack[stack_index] = null as unknown as SmolXMLNode
          if (stack_index === 0) {
            on_node(current_node!)
            current_node = null
            stack_index = -1
          } else {
            stack[stack_index] = null as unknown as SmolXMLNode
            stack_index--
            current_node = stack[stack_index]
          }
        }
      }

      if (i < len && file[i] === '>') {
        i++
      }
    }
  }
}
