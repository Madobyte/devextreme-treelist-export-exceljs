class TreeListHelpers {
  constructor(component, worksheet, options) {
    this.component = component;
    this.worksheet = worksheet;
    this.columns = this.component.getVisibleColumns();

    this.rootValue = this.component.option('rootValue');
    this.parentIdExpr = this.component.option('parentIdExpr');
    this.keyExpr = this.component.option('keyExpr');
    this.dataStructure = this.component.option('dataStructure');

    this.data = this._getData();

    this.worksheet.properties.outlineProperties = {
      summaryBelow: false,
      summaryRight: false,
    };
  }

  _getData() {
    let data = [];
    this.component
      .getDataSource()
      .store()
      .load()
      .done((result) => (data = result));
    if (this.dataStructure === 'plain')
      data = this._convertToHierarchical(data);
    return this._depthDecorator(data);
  }

  // adds the depth for hierarchical data
  _depthDecorator(data, depth = 0) {
    return data.map((node) =>
      Object.assign(node, {
        ...node,
        depth,
        items: this._depthDecorator(node.items || [], depth + 1),
      })
    );
  }

  // converts plain to hierarchical
  _convertToHierarchical(data, id = this.rootValue) {
    return data
      .filter((node) => node[this.parentIdExpr] === id)
      .map((node) => ({
        ...node,
        items: this._convertToHierarchical(data, node[this.keyExpr]),
      }));
  }

  _exportRows(rows) {
    rows.forEach((row) => {
      this._exportRow(row);

      if (this._hasChildren(row)) this._exportRows(row.items);
    });
  }

  _exportRow(row) {
    const insertedRow = this.worksheet.addRow(row);
    insertedRow.outlineLevel = row.depth;
    this.worksheet.getCell(`A${insertedRow.number}`).alignment = {
      indent: row.depth * 2,
    };
  }

  _generateColumns() {
    this.worksheet.columns = this.columns.map(({ caption, dataField }) => ({
      header: caption,
      key: dataField,
    }));
  }

  _hasChildren(row) {
    return row.items && row.items.length > 0;
  }

  _adjustColumnsWidth() {
    this.worksheet.columns.forEach((column) => {
      const lengths = column.values.map((v) => v.toString().length);
      const maxLength = Math.max(
        ...lengths.filter((v) => typeof v === 'number')
      );
      column.width = maxLength;
    });
  }

  export() {
    this._generateColumns();
    this._exportRows(this.data);
    this._adjustColumnsWidth();
  }
}

function exportTreeList({ component, worksheet }) {
  const helpers = new TreeListHelpers(component, worksheet);
  return new Promise((resolve, reject) => {
    helpers.export();
    resolve();
  });
}

module.exports = { exportTreeList };
