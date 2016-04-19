import Ember from 'ember';

export function formatHours(params/*, hash*/) {
  var days = params[0];
  var hours = Math.floor(days*24);
  var minutes = Math.round(((days*24)-Math.floor(days*24))*60);

  if (minutes == 60) {
    hours += 1;
    minutes = 0;
  }

  var formattedHours = ("00"+hours).slice(-2) + ":" + ("00"+minutes).slice(-2);
  return formattedHours;
}

export default Ember.Helper.helper(formatHours);
