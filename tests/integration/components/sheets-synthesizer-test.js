import { moduleForComponent, test } from 'ember-qunit';
import hbs from 'htmlbars-inline-precompile';

moduleForComponent('sheets-synthesizer', 'Integration | Component | sheets synthesizer', {
  integration: true
});

test('it renders', function(assert) {
  // Set any properties with this.set('myProperty', 'value');
  // Handle any actions with this.on('myAction', function(val) { ... });

  this.render(hbs`{{sheets-synthesizer}}`);

  assert.equal(this.$().text().trim(), '');

  // Template block usage:
  this.render(hbs`
    {{#sheets-synthesizer}}
      template block text
    {{/sheets-synthesizer}}
  `);

  assert.equal(this.$().text().trim(), 'template block text');
});
